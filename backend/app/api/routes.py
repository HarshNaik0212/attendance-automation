import base64
import contextlib
import io
import logging
import os
import tempfile
import traceback
from pathlib import Path

from fastapi import APIRouter, BackgroundTasks, File, HTTPException, UploadFile
from fastapi.responses import FileResponse

from app.attendance_automation import AttendanceWorkbookProcessor

logger = logging.getLogger(__name__)

router = APIRouter()


@router.get("/health")
async def health():
    return {"status": "ok"}


@router.post("/process")
async def process_file(background_tasks: BackgroundTasks, file: UploadFile = File(...)):
    # Validate file extension
    filename = file.filename or ""
    if not filename.lower().endswith(".xlsx"):
        raise HTTPException(status_code=400, detail="Only .xlsx files are supported.")

    # Prepare temp files
    suffix = ".xlsx"
    # Preserve original stem for output filename
    original_stem = Path(filename).stem
    output_filename = f"{original_stem}_output.xlsx"

    tmp_input = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp_output = tempfile.NamedTemporaryFile(delete=False, suffix=suffix)
    tmp_input_path = tmp_input.name
    tmp_output_path = tmp_output.name
    tmp_input.close()
    tmp_output.close()

    def cleanup():
        for p in (tmp_input_path, tmp_output_path):
            try:
                if os.path.exists(p):
                    os.unlink(p)
            except Exception:
                pass

    try:
        # Save uploaded file to temp input
        content = await file.read()
        if not content:
            raise HTTPException(status_code=422, detail="Could not read the uploaded file. Please check it's a valid Excel file.")
        with open(tmp_input_path, "wb") as f:
            f.write(content)

        # Capture stdout warnings from processor
        captured = io.StringIO()
        try:
            with contextlib.redirect_stdout(captured):
                processor = AttendanceWorkbookProcessor(tmp_input_path)
                results = processor.process(tmp_output_path)
        except HTTPException:
            raise
        except Exception as e:
            tb = traceback.format_exc()
            logger.error("Processing failed: %s\n%s", e, tb)
            msg = str(e)
            # Detect missing MAIN sheet
            if "MAIN" in msg or "sheet" in msg.lower() and "MAIN" in msg:
                raise HTTPException(status_code=422, detail="Expected a sheet named 'MAIN' in the uploaded file.")
            # Check for openpyxl / zip errors that indicate corrupt file
            lower_msg = msg.lower()
            if any(kw in lower_msg for kw in ["not a valid", "corrupt", "cannot open", "badzip", "invalidfile", "file is not a zip file", "no such file"]):
                raise HTTPException(status_code=422, detail="Could not read the uploaded file. Please check it's a valid Excel file.")
            # More generic: if processor failed to parse dates or no dates found, treat as 422 if it's a user file issue
            # Check if error originates from missing sheet logic
            if "MAIN" in tb:
                raise HTTPException(status_code=422, detail="Expected a sheet named 'MAIN' in the uploaded file.")
            # Fallback: try to distinguish file-structure errors (422) vs unexpected (500)
            # We treat InvalidFileException, BadZipFile, KeyError for sheet as 422
            if isinstance(e, KeyError) and "MAIN" in str(e):
                raise HTTPException(status_code=422, detail="Expected a sheet named 'MAIN' in the uploaded file.")
            # If error message suggests file read failure, use 422
            try:
                # Attempt to detect openpyxl InvalidFileException
                from openpyxl.utils.exceptions import InvalidFileException
                if isinstance(e, InvalidFileException):
                    raise HTTPException(status_code=422, detail="Could not read the uploaded file. Please check it's a valid Excel file.")
            except ImportError:
                pass
            import zipfile
            if isinstance(e, zipfile.BadZipFile):
                raise HTTPException(status_code=422, detail="Could not read the uploaded file. Please check it's a valid Excel file.")
            # For other ValueError/IndexError during parsing, treat as 422 if it clearly is file structure
            # Otherwise 500
            # Heuristic: if results not yet defined, it's likely file structure
            # We'll return 422 for most processing errors to avoid leaking 500 for bad files,
            # but log as error. Use 422 with human-readable message.
            # To respect spec: corrupted/empty -> 422, other unexpected -> 500
            # We'll use 422 if the error looks like file-related, else 500
            # Default to 422 for safety, but if traceback contains our business logic, it's still file data issue
            # Let's return 422 with the original message sanitized, unless it's clearly a code bug
            # For now, return 422 with detail from exception if it's short, else generic
            detail = str(e) if len(str(e)) < 200 else "Could not read the uploaded file. Please check it's a valid Excel file."
            # If detail is empty, use generic
            if not detail.strip():
                detail = "Could not read the uploaded file. Please check it's a valid Excel file."
            # If the original error was very generic, use generic message
            # We avoid leaking raw traceback, but give the message
            raise HTTPException(status_code=422, detail=detail)

        # Ensure output file exists
        if not os.path.exists(tmp_output_path) or os.path.getsize(tmp_output_path) == 0:
            raise HTTPException(status_code=422, detail="Could not read the uploaded file. Please check it's a valid Excel file.")

        employee_count = len(results)
        warnings_text = captured.getvalue().strip()

        headers = {
            "X-Employees-Processed": str(employee_count),
        }
        if warnings_text:
            # Headers must be ASCII-safe — base64 encode
            encoded = base64.b64encode(warnings_text.encode("utf-8")).decode("ascii")
            headers["X-Warnings"] = encoded
            # Expose custom headers for browser CORS
            headers["Access-Control-Expose-Headers"] = "X-Employees-Processed, X-Warnings"

        # Also always expose employee count header
        if "Access-Control-Expose-Headers" not in headers:
            headers["Access-Control-Expose-Headers"] = "X-Employees-Processed, X-Warnings"

        background_tasks.add_task(cleanup)

        return FileResponse(
            path=tmp_output_path,
            filename=output_filename,
            media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            headers=headers,
            background=background_tasks,
        )

    except HTTPException as he:
        cleanup()
        raise he
    except Exception as e:
        cleanup()
        tb = traceback.format_exc()
        logger.error("Unexpected error: %s\n%s", e, tb)
        raise HTTPException(status_code=500, detail="Something went wrong while processing the file.")
