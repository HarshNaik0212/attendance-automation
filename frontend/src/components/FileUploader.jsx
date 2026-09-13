import { useRef, useState } from 'react'

export default function FileUploader({ onFileSelect }) {
  const [dragOver, setDragOver] = useState(false)
  const inputRef = useRef(null)

  const handleDragOver = (e) => {
    e.preventDefault()
    setDragOver(true)
  }

  const handleDragLeave = (e) => {
    e.preventDefault()
    setDragOver(false)
  }

  const handleDrop = (e) => {
    e.preventDefault()
    setDragOver(false)
    const file = e.dataTransfer.files?.[0]
    if (file) onFileSelect(file)
  }

  const handleClick = () => {
    inputRef.current?.click()
  }

  const handleChange = (e) => {
    const file = e.target.files?.[0]
    if (file) onFileSelect(file)
  }

  return (
    <>
      <div
        className={`dropzone ${dragOver ? 'dragover' : ''}`}
        onDragOver={handleDragOver}
        onDragLeave={handleDragLeave}
        onDrop={handleDrop}
        onClick={handleClick}
        role="button"
        tabIndex={0}
        onKeyDown={(e) => { if (e.key === 'Enter' || e.key === ' ') handleClick() }}
      >
        <p>Drag &amp; drop your attendance Excel file here, or click to browse.</p>
        <small>Only .xlsx files are accepted</small>
      </div>
      <input
        ref={inputRef}
        type="file"
        accept=".xlsx"
        style={{ display: 'none' }}
        onChange={handleChange}
      />
    </>
  )
}
