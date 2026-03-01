import { useState, useEffect } from 'react'
import axios from 'axios'
import InsurancePolicyForm from '../../components/InsurancePolicyForm'
import './InsurancePolicies.css'

function InsurancePolicies() {
  const [policies, setPolicies] = useState([])
  const [loading, setLoading] = useState(true)
  const [error, setError] = useState(null)
  
  // Pagination
  const [page, setPage] = useState(1)
  const [limit, setLimit] = useState(10)
  const [total, setTotal] = useState(0)
  const [totalPages, setTotalPages] = useState(0)
  
  // Sorting
  const [sortBy, setSortBy] = useState('id')
  const [sortOrder, setSortOrder] = useState('asc')
  
  // Filtering
  const [filters, setFilters] = useState({
    company: '',
    agent_inner_code: '',
    polis_number: '',
    region: '',
    owner_name: '',
    minPrice: '',
    maxPrice: ''
  })

  // Form state
  const [showForm, setShowForm] = useState(false)

  useEffect(() => {
    fetchPolicies()
  }, [page, limit, sortBy, sortOrder, filters])

  const fetchPolicies = async () => {
    try {
      setLoading(true)
      
      const params = new URLSearchParams({
        page,
        limit,
        sortBy,
        sortOrder,
        ...(filters.company && { company: filters.company }),
        ...(filters.agent_inner_code && { agent_inner_code: filters.agent_inner_code }),
        ...(filters.polis_number && { polis_number: filters.polis_number }),
        ...(filters.region && { region: filters.region }),
        ...(filters.owner_name && { owner_name: filters.owner_name }),
        ...(filters.minPrice && { minPrice: filters.minPrice }),
        ...(filters.maxPrice && { maxPrice: filters.maxPrice })
      })
      
      const response = await axios.get(`/api/insurance-policies?${params}`)
      setPolicies(response.data.data || [])
      setTotal(response.data.pagination.total)
      setTotalPages(response.data.pagination.totalPages)
      setError(null)
    } catch (err) {
      setError(err.message || 'Failed to fetch insurance policies')
      setPolicies([])
    } finally {
      setLoading(false)
    }
  }
  
  const handleFilterChange = (field, value) => {
    setFilters(prev => ({ ...prev, [field]: value }))
    setPage(1)
  }
  
  const handleSortChange = (field) => {
    if (sortBy === field) {
      setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc')
    } else {
      setSortBy(field)
      setSortOrder('asc')
    }
  }
  
  const resetFilters = () => {
    setFilters({
      company: '',
      agent_inner_code: '',
      polis_number: '',
      region: '',
      owner_name: '',
      minPrice: '',
      maxPrice: ''
    })
    setPage(1)
  }

  return (
    <div className="insurance-policies">
      <div className="page-header">
        <h2>Պայմանագրեր</h2>
        <div className="header-buttons">
          <button onClick={fetchPolicies} className="refresh-btn">🔄 Թարմացնել</button>
          <button 
            onClick={() => setShowForm(!showForm)} 
            className="add-policy-btn"
          >
            {showForm ? '✕ Փակել' : '➕ Ավելացնել նոր'}
          </button>
        </div>
      </div>

      {/* Add Policy Form */}
      {showForm && (
        <InsurancePolicyForm 
          onClose={() => setShowForm(false)}
          onSuccess={fetchPolicies}
        />
      )}

      {/* Filters */}
      <div className="filters-section">
        <div className="filters-grid">
          <input
            type="text"
            placeholder="Ընկերություն"
            value={filters.company}
            onChange={(e) => handleFilterChange('company', e.target.value)}
            className="filter-input"
          />
          <input
            type="text"
            placeholder="Գործակալի կոդը"
            value={filters.agent_inner_code}
            onChange={(e) => handleFilterChange('agent_inner_code', e.target.value)}
            className="filter-input"
          />
          <input
            type="text"
            placeholder="Պոլիսի համարը"
            value={filters.polis_number}
            onChange={(e) => handleFilterChange('polis_number', e.target.value)}
            className="filter-input"
          />
          <input
            type="text"
            placeholder="Շրջան"
            value={filters.region}
            onChange={(e) => handleFilterChange('region', e.target.value)}
            className="filter-input"
          />
          <input
            type="text"
            placeholder="Տիրապետի անունը"
            value={filters.owner_name}
            onChange={(e) => handleFilterChange('owner_name', e.target.value)}
            className="filter-input"
          />
          <div className="price-range">
            <input
              type="number"
              placeholder="Նվազ. գինը"
              value={filters.minPrice}
              onChange={(e) => handleFilterChange('minPrice', e.target.value)}
              className="filter-input"
            />
            <input
              type="number"
              placeholder="Մաքս. գինը"
              value={filters.maxPrice}
              onChange={(e) => handleFilterChange('maxPrice', e.target.value)}
              className="filter-input"
            />
          </div>
        </div>
        <button onClick={resetFilters} className="reset-btn">Չեղարկել ֆիլտրերը</button>
      </div>

      {/* Sorting and Pagination Controls */}
      <div className="controls-section">
        <div className="sort-control">
          <label>Տեսակավորեք:</label>
          <select value={sortBy} onChange={(e) => setSortBy(e.target.value)} className="sort-select">
            <option value="id">ID</option>
            <option value="company">Ընկերություն</option>
            <option value="agent_inner_code">Գործակալ</option>
            <option value="polis_number">Պոլիս</option>
            <option value="owner_name">Տիրապետ</option>
            <option value="start_date">Մեկնարկ</option>
            <option value="end_date">Ավարտ</option>
            <option value="region">Շրջան</option>
            <option value="price">Գինը</option>
            <option value="agent_income">Գործակ. եկամուտ</option>
            <option value="income">Եկամուտ</option>
          </select>
          <button 
            onClick={() => setSortOrder(sortOrder === 'asc' ? 'desc' : 'asc')}
            className="sort-order-btn"
          >
            {sortOrder === 'asc' ? '↑' : '↓'}
          </button>
        </div>
        
        <div className="pagination-control">
          <label>Տողեր մեկ էջում:</label>
          <select value={limit} onChange={(e) => { setLimit(parseInt(e.target.value)); setPage(1); }} className="limit-select">
            <option value="5">5</option>
            <option value="10">10</option>
            <option value="25">25</option>
            <option value="50">50</option>
            <option value="100">100</option>
          </select>
        </div>
      </div>

      <div className="page-content">
        {loading && <div className="loading">Բեռնվում է...</div>}
        {error && <div className="error">Սխալ: {error}</div>}
        {!loading && !error && (
          <>
            <div className="table-wrapper">
              <table className="policies-table">
                <thead>
                  <tr>
                    <th onClick={() => handleSortChange('id')} className={sortBy === 'id' ? 'sortable active' : 'sortable'}>ID {sortBy === 'id' && (sortOrder === 'asc' ? '↑' : '↓')}</th>
                    <th onClick={() => handleSortChange('company')} className={sortBy === 'company' ? 'sortable active' : 'sortable'}>Ընկերություն {sortBy === 'company' && (sortOrder === 'asc' ? '↑' : '↓')}</th>
                    <th onClick={() => handleSortChange('polis_number')} className={sortBy === 'polis_number' ? 'sortable active' : 'sortable'}>Պոլիս {sortBy === 'polis_number' && (sortOrder === 'asc' ? '↑' : '↓')}</th>
                    <th onClick={() => handleSortChange('owner_name')} className={sortBy === 'owner_name' ? 'sortable active' : 'sortable'}>Անուն {sortBy === 'owner_name' && (sortOrder === 'asc' ? '↑' : '↓')}</th>
                    <th onClick={() => handleSortChange('agent_inner_code')} className={sortBy === 'agent_inner_code' ? 'sortable active' : 'sortable'}>Գործակալ {sortBy === 'agent_inner_code' && (sortOrder === 'asc' ? '↑' : '↓')}</th>
                    <th onClick={() => handleSortChange('start_date')} className={sortBy === 'start_date' ? 'sortable active' : 'sortable'}>Մեկնարկ {sortBy === 'start_date' && (sortOrder === 'asc' ? '↑' : '↓')}</th>
                    <th onClick={() => handleSortChange('end_date')} className={sortBy === 'end_date' ? 'sortable active' : 'sortable'}>Ավարտ {sortBy === 'end_date' && (sortOrder === 'asc' ? '↑' : '↓')}</th>
                    <th onClick={() => handleSortChange('region')} className={sortBy === 'region' ? 'sortable active' : 'sortable'}>Շրջան {sortBy === 'region' && (sortOrder === 'asc' ? '↑' : '↓')}</th>
                    <th>BM</th>
                    <th>Մոդել</th>
                    <th>Համարակ.</th>
                    <th>HP</th>
                    <th>Ժամանակաշրջան</th>
                    <th onClick={() => handleSortChange('price')} className={sortBy === 'price' ? 'sortable active' : 'sortable'}>Գինը {sortBy === 'price' && (sortOrder === 'asc' ? '↑' : '↓')}</th>
                    <th>Գործակ. %</th>
                    <th>Ընկ. %</th>
                    <th onClick={() => handleSortChange('agent_income')} className={sortBy === 'agent_income' ? 'sortable active' : 'sortable'}>Գործակ. եկամուտը {sortBy === 'agent_income' && (sortOrder === 'asc' ? '↑' : '↓')}</th>
                    <th onClick={() => handleSortChange('income')} className={sortBy === 'income' ? 'sortable active' : 'sortable'}>Եկամուտը {sortBy === 'income' && (sortOrder === 'asc' ? '↑' : '↓')}</th>
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
                        <td>{item.agent_inner_code}</td>
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

export default InsurancePolicies
