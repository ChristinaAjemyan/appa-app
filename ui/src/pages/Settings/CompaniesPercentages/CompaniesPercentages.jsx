import { useState, useEffect } from 'react'
import axios from 'axios'
import Modal from '../../../components/Modal'
import PercentageForm from '../../../components/PercentageForm'
import './CompaniesPercentages.css'

function CompaniesPercentages() {
  const [percentages, setPercentages] = useState([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(null)
  const [showModal, setShowModal] = useState(false)
  const [editingItem, setEditingItem] = useState(null)
  const [isSaving, setIsSaving] = useState(false)

  // Pagination
  const [page, setPage] = useState(1)
  const [limit, setLimit] = useState(10)
  const [total, setTotal] = useState(0)
  const [totalPages, setTotalPages] = useState(0)

  useEffect(() => {
    fetchPercentages()
  }, [page, limit])

  const fetchPercentages = async () => {
    try {
      setLoading(true)
      const response = await axios.get(`/api/companies-percentage?page=${page}&limit=${limit}`)
      setPercentages(response.data.data || [])
      setTotal(response.data.pagination.total)
      setTotalPages(response.data.pagination.totalPages)
      setError(null)
    } catch (err) {
      setError(err.message || 'Failed to fetch companies percentages')
      setPercentages([])
    } finally {
      setLoading(false)
    }
  }

  const handleAddNew = () => {
    setEditingItem(null)
    setShowModal(true)
  }

  const handleEdit = (item) => {
    setEditingItem(item)
    setShowModal(true)
  }

  const handleDelete = async (id) => {
    if (!window.confirm('Վստահ եք?')) return
    
    try {
      await axios.delete(`/api/companies-percentage/${id}`)
      await fetchPercentages()
    } catch (err) {
      alert('Error deleting record: ' + err.message)
    }
  }

  const handleSave = async (formData) => {
    try {
      setIsSaving(true)
      if (editingItem?.id) {
        await axios.put(`/api/companies-percentage/${editingItem.id}`, formData)
      } else {
        await axios.post('/api/companies-percentage', formData)
      }
      setShowModal(false)
      await fetchPercentages()
    } catch (err) {
      alert('Error saving record: ' + err.message)
    } finally {
      setIsSaving(false)
    }
  }

  return (
    <div className="companies-percentages">
      <div className="page-header">
        <h2>Ընկերությունների տոկոսներ</h2>
        <div className="header-actions">
          <button onClick={fetchPercentages} className="refresh-btn">🔄 Թարմացնել</button>
          <button onClick={handleAddNew} className="add-btn">+ Ավելացնել</button>
        </div>
      </div>

      <div className="page-content">
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
              <table className="percentages-table">
                <thead>
                  <tr>
                    <th>ID</th>
                    <th>Ընկերություն</th>
                    <th>Գործակալ Կոդ (In)</th>
                    <th>Գործակալ Կոդ (Not)</th>
                    <th>Շրջան (In)</th>
                    <th>Շրջան (Not)</th>
                    <th>BM Min</th>
                    <th>BM Max</th>
                    <th>BM Exact</th>
                    <th>Brand In</th>
                    <th>HP Min</th>
                    <th>HP Max</th>
                    <th>Ժամանակաշրջան</th>
                    <th>Տոկոս</th>
                    <th>Գործողություն</th>
                  </tr>
                </thead>
                <tbody>
                  {percentages.length > 0 ? (
                    percentages.map((item) => (
                      <tr key={item.id}>
                        <td>{item.id}</td>
                        <td>{item.company}</td>
                        <td>{item.agent_code_in}</td>
                        <td>{item.agent_code_not}</td>
                        <td>{item.region_in}</td>
                        <td>{item.region_not}</td>
                        <td>{item.bm_min}</td>
                        <td>{item.bm_max}</td>
                        <td>{item.bm_exact}</td>
                        <td>{item.brand_in}</td>
                        <td>{item.hp_min}</td>
                        <td>{item.hp_max}</td>
                        <td>{item.period}</td>
                        <td className="percentage">{item.percentage}%</td>
                        <td className="actions-cell">
                          <button className="btn-edit" onClick={() => handleEdit(item)}>Խմբագր</button>
                          <button className="btn-delete" onClick={() => handleDelete(item.id)}>Ջնջել</button>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan="15" className="empty-state">Տվյալներ չկան</td>
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

      <Modal
        isOpen={showModal}
        title={editingItem ? 'Խմբագրել Տոկոսը' : 'Նոր Տոկոս'}
        onClose={() => setShowModal(false)}
      >
        <PercentageForm
          item={editingItem}
          onSave={handleSave}
          onCancel={() => setShowModal(false)}
          isLoading={isSaving}
        />
      </Modal>
    </div>
  )
}

export default CompaniesPercentages
