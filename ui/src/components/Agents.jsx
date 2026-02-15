import { useState, useEffect } from 'react'
import axios from 'axios'
import './Agents.css'

function Agents() {
  const [agents, setAgents] = useState([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(null)

  useEffect(() => {
    fetchAgents()
  }, [])

  const fetchAgents = async () => {
    try {
      setLoading(true)
      const response = await axios.get('/api/agents')
      setAgents(response.data || [])
      setError(null)
    } catch (err) {
      setError(err.message || 'Failed to fetch agents')
      setAgents([])
    } finally {
      setLoading(false)
    }
  }

  return (
    <div className="agents">
      <div className="agents-header">
        <h2>Գործակալներ</h2>
        <button onClick={fetchAgents} className="refresh-btn">🔄 Թարմացնել</button>
      </div>

      <div className="agents-content">
        {loading && <div className="loading">Բեռնվում է...</div>}
        {error && <div className="error">Սխալ: {error}</div>}
        {!loading && !error && (
          <div className="table-wrapper">
            <table className="agents-table">
              <thead>
                <tr>
                  <th>ID</th>
                  <th>Կոդ</th>
                  <th>Անուն</th>
                  <th>Հեռ․ համար</th>
                </tr>
              </thead>
              <tbody>
                {agents.length > 0 ? (
                  agents.map((agent) => (
                    <tr key={agent.id}>
                      <td>{agent.id}</td>
                      <td>{agent.code}</td>
                      <td>{agent.name}</td>
                      <td>{agent.phone_number || '-'}</td>
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

export default Agents
