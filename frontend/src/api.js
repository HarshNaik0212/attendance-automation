export async function processFile(file) {
  const formData = new FormData()
  formData.append('file', file)

  const response = await fetch('/api/process', {
    method: 'POST',
    body: formData,
  })

  if (!response.ok) {
    let detail = 'Something went wrong while processing the file.'
    try {
      const data = await response.json()
      if (data && data.detail) detail = data.detail
    } catch {
      // response may not be JSON
    }
    const error = new Error(detail)
    error.status = response.status
    throw error
  }

  const blob = await response.blob()
  const contentDisposition = response.headers.get('Content-Disposition') || ''
  let filename = 'output.xlsx'
  const match = contentDisposition.match(/filename="?([^"]+)"?/)
  if (match) filename = match[1]

  const employeesHeader = response.headers.get('X-Employees-Processed')
  const warningsHeader = response.headers.get('X-Warnings')

  let employeeCount = null
  if (employeesHeader) employeeCount = parseInt(employeesHeader, 10)

  let warnings = null
  if (warningsHeader) {
    try {
      warnings = atob(warningsHeader)
    } catch {
      warnings = warningsHeader
    }
  }

  return { blob, filename, employeeCount, warnings }
}
