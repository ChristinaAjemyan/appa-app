import { useState, useEffect } from 'react'
import axios from 'axios'
import './Dashboard.css'

function Dashboard() {
  const [data, setData] = useState(null)
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(null)
  const [file, setFile] = useState(null)
  const [year, setYear] = useState(new Date().getFullYear())
  const [month, setMonth] = useState(new Date().getMonth() + 1)
  const [uploadLoading, setUploadLoading] = useState(false)
  const [uploadError, setUploadError] = useState(null)
  const [uploadSuccess, setUploadSuccess] = useState(null)

  useEffect(() => {
    fetchData()
  }, [])

  const fetchData = async () => {
    try {
      setLoading(true)
      const response = await axios.get('/api/health')
      setData(response.data)
      setError(null)
    } catch (err) {
      setError(err.message || 'Failed to fetch data')
      setData(null)
    } finally {
      setLoading(false)
    }
  }

  const handleFileChange = (e) => {
    setFile(e.target.files[0])
    setUploadError(null)
    setUploadSuccess(null)
  }

  const handleSubmit = async (e) => {
    e.preventDefault()
    
    if (!file) {
      setUploadError('Please select a file')
      return
    }

    if (!file.name.endsWith('.csv')) {
      setUploadError('Please select a CSV file')
      return
    }

    const formData = new FormData()
    formData.append('file', file)
    formData.append('year', year)
    formData.append('month', month)

    try {
      setUploadLoading(true)
      setUploadError(null)
      const response = await axios.post('/api/calculate', formData, {
        headers: {
          'Content-Type': 'multipart/form-data',
        },
      })
      setUploadSuccess(`Successfully imported ${response.data.stats.inserted} records`)
      setFile(null)
      document.querySelector('input[type="file"]').value = ''
    } catch (err) {
      setUploadError(err.response?.data?.error || err.message || 'Upload failed')
    } finally {
      setUploadLoading(false)
    }
  }

  const currentYear = new Date().getFullYear()
  const years = Array.from({ length: 10 }, (_, i) => currentYear - i)
  const months = [
    { value: 1, label: 'Հունվար' },
    { value: 2, label: 'Փետրվար' },
    { value: 3, label: 'Մարտ' },
    { value: 4, label: 'Ապրիլ' },
    { value: 5, label: 'Մայիս' },
    { value: 6, label: 'Հունիս' },
    { value: 7, label: 'Հուլիս' },
    { value: 8, label: 'Օգոստոս' },
    { value: 9, label: 'Սեպտեմբեր' },
    { value: 10, label: 'Հոկտեմբեր' },
    { value: 11, label: 'Նոյեմբեր' },
    { value: 12, label: 'Դեկեմբեր' },
  ]

  return (
    <div className="dashboard">
      <div className="dashboard-header">
        <h2>Հաշվարկ</h2>
        <button onClick={fetchData} className="refresh-btn">🔄 Թարմացնել</button>
      </div>

      <div className="dashboard-content">
        {loading && <div className="loading">Բեռնվում է...</div>}
        {error && <div className="error">Սխալ: {error}</div>}
        {data && (
          <div className="card">
            <h3>Սերվերի կարգավիճակ</h3>
            <p className="status-text">{data.status}</p>
          </div>
        )}

        <div className="card upload-card">
          <h3>CSV ֆայլ վերբեռնել</h3>
          <form onSubmit={handleSubmit} className="upload-form">
            <div className="form-group">
              <label htmlFor="file">Ընտրել ֆայլ:</label>
              <input
                type="file"
                id="file"
                accept=".csv"
                onChange={handleFileChange}
                disabled={uploadLoading}
              />
              {file && <p className="file-info">✓ {file.name}</p>}
            </div>

            <div className="form-row">
              <div className="form-group">
                <label htmlFor="year">Տարի:</label>
                <select
                  id="year"
                  value={year}
                  onChange={(e) => setYear(Number(e.target.value))}
                  disabled={uploadLoading}
                >
                  {years.map((y) => (
                    <option key={y} value={y}>
                      {y}
                    </option>
                  ))}
                </select>
              </div>

              <div className="form-group">
                <label htmlFor="month">Ամիս:</label>
                <select
                  id="month"
                  value={month}
                  onChange={(e) => setMonth(Number(e.target.value))}
                  disabled={uploadLoading}
                >
                  {months.map((m) => (
                    <option key={m.value} value={m.value}>
                      {m.label}
                    </option>
                  ))}
                </select>
              </div>
            </div>

            <button
              type="submit"
              className="submit-btn"
              disabled={uploadLoading || !file}
            >
              {uploadLoading ? 'Վերբեռնվում է...' : 'Հաշվարկել'}
            </button>
          </form>

          {uploadError && <div className="error">{uploadError}</div>}
          {uploadSuccess && <div className="success">{uploadSuccess}</div>}
        </div>
      </div>
    </div>
  )
}

export default Dashboard
