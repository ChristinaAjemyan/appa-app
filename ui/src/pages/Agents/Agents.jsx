import { useState, useEffect } from 'react'
import axios from 'axios'
import Modal from '../../components/Modal'
import AgentForm from '../../components/AgentForm'
import './Agents.css'

function Agents() {
  const [agents, setAgents] = useState([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(null)
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [selectedAgent, setSelectedAgent] = useState(null)
  const [isFormLoading, setIsFormLoading] = useState(false)

  // Pagination
  const [page, setPage] = useState(1)
  const [limit, setLimit] = useState(10)
  const [total, setTotal] = useState(0)
  const [totalPages, setTotalPages] = useState(0)

  useEffect(() => {
    fetchAgents()
  }, [page, limit])

  const fetchAgents = async () => {
    try {
      setLoading(true)
      const response = await axios.get(`/api/agents?page=${page}&limit=${limit}`)
      setAgents(response.data.data || [])
      setTotal(response.data.pagination.total)
      setTotalPages(response.data.pagination.totalPages)
      setError(null)
    } catch (err) {
      setError(err.message || 'Failed to fetch agents')
      setAgents([])
    } finally {
      setLoading(false)
    }
  }

  const handleAddAgent = () => {
    setSelectedAgent(null)
    setIsModalOpen(true)
  }

  const handleEditAgent = (agent) => {
    setSelectedAgent(agent)
    setIsModalOpen(true)
  }

  const handleDeleteAgent = async (id) => {
    if (window.confirm('Հաստատել ջնջումը?')) {
      try {
        await axios.delete(`/api/agents/${id}`)
        fetchAgents()
      } catch (err) {
        alert('Error deleting agent: ' + err.message)
      }
    }
  }

  const handleSaveAgent = async (formData) => {
    try {
      setIsFormLoading(true)
      if (selectedAgent && selectedAgent.id) {
        await axios.put(`/api/agents/${selectedAgent.id}`, formData)
      } else {
        await axios.post('/api/agents', formData)
      }
      setIsModalOpen(false)
      fetchAgents()
    } catch (err) {
      alert('Error saving agent: ' + err.message)
    } finally {
      setIsFormLoading(false)
    }
  }

  const handleCloseModal = () => {
    setIsModalOpen(false)
    setSelectedAgent(null)
  }

  return (
    <div className="agents">
      <div className="agents-header">
        <h2>Գործակալներ</h2>
        <div className="header-buttons">
          <button onClick={fetchAgents} className="refresh-btn">🔄 Թարմացնել</button>
          <button onClick={handleAddAgent} className="add-btn">➕ Ավելացնել</button>
        </div>
      </div>

      <Modal
        isOpen={isModalOpen}
        title={selectedAgent ? 'Խմբագրել Գործակալ' : 'Ավելացնել Գործակալ'}
        onClose={handleCloseModal}
      >
        <AgentForm
          agent={selectedAgent}
          onSave={handleSaveAgent}
          onCancel={handleCloseModal}
          isLoading={isFormLoading}
        />
      </Modal>

      <div className="agents-content">
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
              <table className="agents-table">
                <thead>
                  <tr>
                    <th>ID</th>
                    <th>Կոդ</th>
                    <th>Անուն</th>
                    {/* <th>Հեռ․ համար</th> */}
                    <th>NAIRI</th>
                    <th>REGO</th>
                    <th>ARMENIA</th>
                    <th>SIL</th>
                    <th>INGO</th>
                    <th>LIGA</th>
                    <th>Գործողություններ</th>
                  </tr>
                </thead>
                <tbody>
                  {agents.length > 0 ? (
                    agents.map((agent) => (
                      <tr key={agent.id}>
                        <td>{agent.id}</td>
                        <td>{agent.code}</td>
                        <td>{agent.name}</td>
                        {/* <td>{agent.phone_number || '-'}</td> */}
                        <td>{agent.nairi || '-'}</td>
                        <td>{agent.rego || '-'}</td>
                        <td>{agent.armenia || '-'}</td>
                        <td>{agent.sil || '-'}</td>
                        <td>{agent.ingo || '-'}</td>
                        <td>{agent.liga || '-'}</td>
                        <td className="actions">
                          <button
                            className="btn-edit"
                            onClick={() => handleEditAgent(agent)}
                            title="Խմբագրել"
                          >
                            ✎
                          </button>
                          <button
                            className="btn-delete"
                            onClick={() => handleDeleteAgent(agent.id)}
                            title="Ջնջել"
                          >
                            ✕
                          </button>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="11" className="empty-state">Տվյալներ չկան</td>
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

export default Agents
