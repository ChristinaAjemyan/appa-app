import { useState, useEffect } from 'react'
import axios from 'axios'
import './Companies.css'

function Companies() {
  const [companies, setCompanies] = useState([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(null)

  // Pagination
  const [page, setPage] = useState(1)
  const [limit, setLimit] = useState(10)
  const [total, setTotal] = useState(0)
  const [totalPages, setTotalPages] = useState(0)

  useEffect(() => {
    fetchCompanies()
  }, [page, limit])

  const fetchCompanies = async () => {
    try {
      setLoading(true)
      const response = await axios.get(`/api/company?page=${page}&limit=${limit}`)
      setCompanies(response.data.data || [])
      setTotal(response.data.pagination.total)
      setTotalPages(response.data.pagination.totalPages)
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
          <>
            <div className="pagination-control">
              <label>Տողեր մեկ էջում:</label>
              <select 
                value={limit} 
                onChange={(e) => { setLimit(parseInt(e.target.value)); setPage(1); }} 
                className="limit-select"
              >
                <option value="5">5</option>
                <option value="10">10</option>
                <option value="25">25</option>
                <option value="50">50</option>
              </select>
            </div>
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
            
            {/* Pagination */}
            <div className="pagination">
              <button 
                onClick={() => setPage(1)} 
                disabled={page === 1}
                className="pagination-btn"
              >
                Առաջին
              </button>
              <button 
                onClick={() => setPage(Math.max(1, page - 1))} 
                disabled={page === 1}
                className="pagination-btn"
              >
                Նախորդ
              </button>
              
              <div className="pagination-info">
                Էջ {page} / {totalPages} (Ընդամենը {total} միավոր)
              </div>
              
              <button 
                onClick={() => setPage(Math.min(totalPages, page + 1))} 
                disabled={page === totalPages}
                className="pagination-btn"
              >
                Հաջորդ
              </button>
              <button 
                onClick={() => setPage(totalPages)} 
                disabled={page === totalPages}
                className="pagination-btn"
              >
                Վերջին
              </button>
            </div>
          </>
        )}
      </div>
    </div>
  )
}

export default Companies
