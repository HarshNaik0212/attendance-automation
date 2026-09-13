import { useState } from 'react'
import FileUploader from './components/FileUploader'
import { ProcessingView, SuccessView, ErrorView } from './components/ProcessingStatus'
import { processFile } from './api'
import './App.css'

function formatFileSize(bytes) {
  if (bytes < 1024) return `${bytes} B`
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`
  return `${(bytes / (1024 * 1024)).toFixed(1)} MB`
}

export default function App() {
  const [status, setStatus] = useState('idle') // idle | selected | processing | success | error
  const [selectedFile, setSelectedFile] = useState(null)
  const [resultBlob, setResultBlob] = useState(null)
  const [resultFilename, setResultFilename] = useState('')
  const [employeeCount, setEmployeeCount] = useState(null)
  const [warnings, setWarnings] = useState(null)
  const [errorMessage, setErrorMessage] = useState('')

  const handleFileSelect = (file) => {
    if (!file.name.toLowerCase().endsWith('.xlsx')) {
      setErrorMessage('Only .xlsx files are supported.')
      setStatus('error')
      return
    }
    setSelectedFile(file)
    setStatus('selected')
    setErrorMessage('')
  }

  const handleProcess = async () => {
    if (!selectedFile) return
    setStatus('processing')
    setErrorMessage('')
    try {
      const { blob, filename, employeeCount: count, warnings: w } = await processFile(selectedFile)
      setResultBlob(blob)
      setResultFilename(filename)
      setEmployeeCount(count)
      setWarnings(w)
      setStatus('success')
    } catch (err) {
      setErrorMessage(err.message || 'Something went wrong while processing the file.')
      setStatus('error')
    }
  }

  const handleDownload = () => {
    if (!resultBlob) return
    const url = URL.createObjectURL(resultBlob)
    const a = document.createElement('a')
    a.href = url
    a.download = resultFilename || 'output.xlsx'
    document.body.appendChild(a)
    a.click()
    a.remove()
    URL.revokeObjectURL(url)
  }

  const handleReset = () => {
    // Revoke old blob URL if needed and reset all state
    setStatus('idle')
    setSelectedFile(null)
    setResultBlob(null)
    setResultFilename('')
    setEmployeeCount(null)
    setWarnings(null)
    setErrorMessage('')
  }

  return (
    <div className="container">
      <div className="card">
        <h1>Attendance Automation</h1>
        <p className="subtitle">Upload your attendance Excel file and get the processed result.</p>

        {status === 'idle' && (
          <FileUploader onFileSelect={handleFileSelect} />
        )}

        {status === 'selected' && selectedFile && (
          <>
            <div className="file-info">
              <div>
                <div className="name">{selectedFile.name}</div>
                <div className="size">{formatFileSize(selectedFile.size)}</div>
              </div>
            </div>
            <button className="btn btn-primary" onClick={handleProcess}>
              Process File
            </button>
            <div style={{ marginTop: '10px' }}>
              <button className="btn btn-secondary" onClick={handleReset}>
                Choose a different file
              </button>
            </div>
          </>
        )}

        {status === 'processing' && <ProcessingView />}

        {status === 'success' && (
          <SuccessView
            employeeCount={employeeCount}
            warnings={warnings}
            onDownload={handleDownload}
            onReset={handleReset}
          />
        )}

        {status === 'error' && (
          <ErrorView message={errorMessage} onReset={handleReset} />
        )}
      </div>
    </div>
  )
}
