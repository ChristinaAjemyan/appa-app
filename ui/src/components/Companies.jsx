import { useState, useEffect } from 'react'
import axios from 'axios'
import './Companies.css'

function Companies() {
  const [companies, setCompanies] = useState([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(null)

  useEffect(() => {
    fetchCompanies()
  }, [])

  const fetchCompanies = async () => {
    try {
      setLoading(true)
      const response = await axios.get('/api/company')
      setCompanies(response.data || [])
      setError(null)
    } catch (err) {
      setError(err.message || 'Failed to fetch companies')
      setCompanies([])
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="companies">
      <div className="companies-header">
        <h3>Ընկերություններ</h3>
        <button onClick={fetchCompanies} className="refresh-btn">🔄 Թարմացնել</button>
      </div>

      <div className="companies-content">
        {loading && <div className="loading">Բեռնվում է...</div>}
        {error && <div className="error">Սխալ: {error}</div>}
        {!loading && !error && (
          <div className="table-wrapper">
            <table className="companies-table">
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Անուն</th>
                  <th>Ընկերության %</th>
                  <th>Գործակալի %</th>
                </tr>
              </thead>
              <tbody>
                {companies.length > 0 ? (
                  companies.map((company) => (
                    <tr key={company.id}>
                      <td>{company.id}</td>
                      <td>{company.name}</td>
                      <td className="percentage">{company.company_percent}%</td>
                      <td className="percentage">{company.agent_percent}%</td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan="4" className="empty-state">Տվյալներ չկան</td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>
        )}
      </div>
    </div>
  )
}

export default Companies
