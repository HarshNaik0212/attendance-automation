export function ProcessingView() {
  return (
    <>
      <div className="spinner" aria-label="loading" />
      <p className="processing-text">Processing your file...</p>
    </>
  )
}

export function SuccessView({ employeeCount, warnings, onDownload, onReset }) {
  return (
    <>
      <div className="success-box">
        ✅ Done — processed {employeeCount ?? '?'} employees.
      </div>
      {warnings && (
        <div className="warning-box">
          <strong>⚠️ Warnings — unrecognized statuses:</strong>
          {warnings}
        </div>
      )}
      <button className="btn btn-download" onClick={onDownload}>
        ⬇ Download Result
      </button>
      <div style={{ marginTop: '12px' }}>
        <button className="btn btn-secondary" onClick={onReset}>
          Process another file
        </button>
      </div>
    </>
  )
}

export function ErrorView({ message, onReset }) {
  return (
    <>
      <div className="error-box">
        <strong>Error:</strong> {message}
      </div>
      <button className="btn btn-primary" onClick={onReset}>
        Try again
      </button>
    </>
  )
}
