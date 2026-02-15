import { useState, useEffect } from 'react'
import axios from 'axios'
import './InsurancePolicies.css'

function InsurancePolicies() {
  const [policies, setPolicies] = useState([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(null)

  useEffect(() => {
    fetchPolicies()
  }, [])

  const fetchPolicies = async () => {
    try {
      setLoading(true)
      const response = await axios.get('/api/insurance-policies')
      setPolicies(response.data || [])
      setError(null)
    } catch (err) {
      setError(err.message || 'Failed to fetch insurance policies')
      setPolicies([])
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="insurance-policies">
      <div className="page-header">
        <h2>Պայմանագրեր</h2>
        <button onClick={fetchPolicies} className="refresh-btn">🔄 Թարմացնել</button>
      </div>

      <div className="page-content">
        {loading && <div className="loading">Բեռնվում է...</div>}
        {error && <div className="error">Սխալ: {error}</div>}
        {!loading && !error && (
          <div className="table-wrapper">
            <table className="policies-table">
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Ընկերություն</th>
                  <th>Պոլիս</th>
                  <th>Մերժ. տերը</th>
                  <th>Գործակալ</th>
                  <th>Մեկնարկ</th>
                  <th>Ավարտ</th>
                  <th>Շրջան</th>
                  <th>BM</th>
                  <th>Մոդել</th>
                  <th>Համարակ.</th>
                  <th>HP</th>
                  <th>Ժամանակաշրջան</th>
                  <th>Գինը</th>
                  <th>Գործակ. %</th>
                  <th>Ընկ. %</th>
                  <th>Գործակ. եկամուտը</th>
                  <th>Եկամուտը</th>
                </tr>
              </thead>
              <tbody>
                {policies.length > 0 ? (
                  policies.map((item) => (
                    <tr key={item.id}>
                      <td>{item.id}</td>
                      <td>{item.company}</td>
                      <td>{item.polis_number}</td>
                      <td>{item.owner_name}</td>
                      <td>{item.agent_name}</td>
                      <td>{item.start_date}</td>
                      <td>{item.end_date}</td>
                      <td>{item.region}</td>
                      <td>{item.bm_class}</td>
                      <td>{item.car_model}</td>
                      <td>{item.car_number}</td>
                      <td>{item.hp}</td>
                      <td>{item.period}</td>
                      <td>{item.price}</td>
                      <td className="percentage">{item.agent_percent}%</td>
                      <td className="percentage">{item.company_percent}%</td>
                      <td className="income">{item.agent_income}</td>
                      <td className="income">{item.income}</td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan="18" className="empty-state">Տվյալներ չկան</td>
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

export default InsurancePolicies
