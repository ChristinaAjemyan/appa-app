import { useState, useEffect } from 'react'
import axios from 'axios'
import './InsurancePolicyForm.css'

function InsurancePolicyForm({ onClose, onSuccess }) {
  const [formData, setFormData] = useState({
    company: '',
    agent_company_code: '',
    agent_inner_code: '',
    agent_name: '',
    polis_number: '',
    owner_name: '',
    start_date: '',
    end_date: '',
    region: '',
    phone_number: '',
    bm_class: '',
    car_model: '',
    car_number: '',
    hp: '',
    period: '',
    info: '',
    price: ''
  })
  const [submitting, setSubmitting] = useState(false)
  const [message, setMessage] = useState(null)
  const [errors, setErrors] = useState({})
  const [companies, setCompanies] = useState([])
  const [companiesLoading, setCompaniesLoading] = useState(true)
  const [agents, setAgents] = useState([])
  const [agentsLoading, setAgentsLoading] = useState(true)
  const [regions, setRegions] = useState([])
  const [regionsLoading, setRegionsLoading] = useState(true)

  useEffect(() => {
    const fetchCompanies = async () => {
      try {
        const response = await axios.get('/api/company')
        setCompanies(response.data.data || response.data)
        setCompaniesLoading(false)
      } catch (err) {
        console.error('Error fetching companies:', err)
        setCompaniesLoading(false)
      }
    }
    const fetchAgents = async () => {
      try {
        const response = await axios.get('/api/agents')
        setAgents(response.data.data || response.data)
        setAgentsLoading(false)
      } catch (err) {
        console.error('Error fetching agents:', err)
        setAgentsLoading(false)
      }
    }
    const fetchRegions = async () => {
      try {
        const response = await axios.get('/api/regions')
        setRegions(response.data.data || response.data)
        setRegionsLoading(false)
      } catch (err) {
        console.error('Error fetching regions:', err)
        setRegionsLoading(false)
      }
    }
    fetchCompanies()
    fetchAgents()
    fetchRegions()
  }, [])

  const handleFormChange = (field, value) => {
    if (field === 'agent_name') {
      const selectedAgent = agents.find(agent => agent.name === value)
      if (selectedAgent) {
        setFormData(prev => ({
          ...prev,
          agent_name: value,
          agent_inner_code: selectedAgent.code
        }))
        return
      }
    }
    setFormData(prev => ({ ...prev, [field]: value }))
    // Clear the error for this field when user modifies it
    if (errors[field]) {
      setErrors(prev => ({ ...prev, [field]: '' }))
    }
  }

  const validateDates = () => {
    const newErrors = {}
    const { start_date, end_date } = formData

    // Check if one date is provided without the other
    if ((start_date && !end_date) || (!start_date && end_date)) {
      if (!start_date && end_date) {
        newErrors.start_date = 'Մեկնարկի ամսաթիվը պարտադիր է, եթե ավարտի ամսաթիվ է նշված'
      }
      if (start_date && !end_date) {
        newErrors.end_date = 'Ավարտի ամսաթիվը պարտադիր է, եթե մեկնարկի ամսաթիվ է նշված'
      }
    }

    // Check if end date is after start date
    if (start_date && end_date && new Date(start_date) >= new Date(end_date)) {
      newErrors.end_date = 'Ավարտի ամսաթիվը պետք է լինի մեկնարկի ամսաթիվից հետո'
    }

    return newErrors
  }

  const handleFormSubmit = async (e) => {
    e.preventDefault()
    
    // Validate dates
    const dateErrors = validateDates()
    if (Object.keys(dateErrors).length > 0) {
      setErrors(dateErrors)
      return
    }

    try {
      setSubmitting(true)
      setMessage(null)

      // Prepare data for submission
      const submitData = {
        ...formData,
        price: formData.price ? parseFloat(formData.price) : null,
        income: formData.price ? parseFloat(formData.price) - parseFloat(formData.discount) : null
      }
      console.log(submitData);
      
      // await axios.post('/api/insurance-policies', submitData)
      
      setMessage({ type: 'success', text: 'Պայմանագիր հաջողությամբ ավելացվել է' })
      
      // Reset form
      resetForm()
      
      // Close form after 2 seconds
      setTimeout(() => {
        onSuccess()
        onClose()
      }, 2000)
    } catch (err) {
      setMessage({ 
        type: 'error', 
        text: err.response?.data?.error || 'Սխալ պայմանագիր ավելացնելիս' 
      })
    } finally {
      setSubmitting(false)
    }
  }

  const resetForm = () => {
    setFormData({
      company: '',
      agent_company_code: '',
      agent_inner_code: '',
      agent_name: '',
      polis_number: '',
      owner_name: '',
      start_date: '',
      end_date: '',
      region: '',
      phone_number: '',
      bm_class: '',
      car_model: '',
      car_number: '',
      hp: '',
      period: '',
      info: '',
      price: ''
    })
    setMessage(null)
    setErrors({})
  }

  return (
    <div className="form-section">
      <h3>Նոր պայմանագիր ավելացնել</h3>
      <form onSubmit={handleFormSubmit} className="insurance-form">
        <div className="form-grid">
          <div className="form-group">
            <label>Ընկերություն *</label>
            <select
              value={formData.company}
              onChange={(e) => handleFormChange('company', e.target.value)}
              className="form-input"
              required
              disabled={companiesLoading}
            >
              <option value="">
                {companiesLoading ? 'Բեռնվում է...' : 'Ընտրել'}
              </option>
              {companies?.map((company) => (
                <option key={company.id} value={company.name}>
                  {company.name}
                </option>
              ))}
            </select>
          </div>

          {/* <div className="form-group">
            <label>Գործակալի ընկերության կոդ</label>
            <input
              type="text"
              value={formData.agent_company_code}
              onChange={(e) => handleFormChange('agent_company_code', e.target.value)}
              className="form-input"
            />
          </div>

          <div className="form-group">
            <label>Գործակալի ներքին կոդ</label>
            <input
              type="text"
              value={formData.agent_inner_code}
              onChange={(e) => handleFormChange('agent_inner_code', e.target.value)}
              className="form-input"
            />
          </div> */}

          <div className="form-group">
            <label>Գործակալ</label>
            <select
              value={formData.agent_name}
              onChange={(e) => handleFormChange('agent_name', e.target.value)}
              className="form-input"
              disabled={agentsLoading}
            >
              <option value="">
                {agentsLoading ? 'Բեռնվում է...' : 'Ընտրել'}
              </option>
              {agents?.map((agent) => (
                <option key={agent.id} value={agent.name}>
                  {agent.name}
                </option>
              ))}
            </select>
          </div>

          <div className="form-group">
            <label>Պոլիսի համար *</label>
            <input
              type="text"
              value={formData.polis_number}
              onChange={(e) => handleFormChange('polis_number', e.target.value)}
              className="form-input"
              required
            />
          </div>

          <div className="form-group">
            <label>Ապահովադիր *</label>
            <input
              type="text"
              value={formData.owner_name}
              onChange={(e) => handleFormChange('owner_name', e.target.value)}
              className="form-input"
              required
            />
          </div>

          <div className="form-group">
            <label>Մեկնարկի ամսաթիվ</label>
            <input
              type="date"
              value={formData.start_date}
              onChange={(e) => handleFormChange('start_date', e.target.value)}
              className={`form-input ${errors.start_date ? 'input-error' : ''}`}
            />
            {errors.start_date && (
              <span className="error-message">{errors.start_date}</span>
            )}
          </div>

          <div className="form-group">
            <label>Ավարտի ամսաթիվ</label>
            <input
              type="date"
              value={formData.end_date}
              onChange={(e) => handleFormChange('end_date', e.target.value)}
              className={`form-input ${errors.end_date ? 'input-error' : ''}`}
            />
            {errors.end_date && (
              <span className="error-message">{errors.end_date}</span>
            )}
          </div>

          <div className="form-group">
            <label>Մարզ</label>
            <select
              value={formData.region}
              onChange={(e) => handleFormChange('region', e.target.value)}
              className="form-input"
              disabled={regionsLoading}
            >
              <option value="">
                {regionsLoading ? 'Բեռնվում է...' : 'Ընտրել'}
              </option>
              {regions?.map((region) => (
                <option key={region.id} value={region.name}>
                  {region.name}
                </option>
              ))}
            </select>
          </div>

          <div className="form-group">
            <label>Հեռախոսի համար</label>
            <input
              type="text"
              value={formData.phone_number}
              onChange={(e) => handleFormChange('phone_number', e.target.value)}
              className="form-input"
            />
          </div>

          <div className="form-group">
            <label>BM դաս</label>
            <input
              type="text"
              value={formData.bm_class}
              onChange={(e) => handleFormChange('bm_class', e.target.value)}
              className="form-input"
            />
          </div>

          <div className="form-group">
            <label>Ավտոմեքենայի մոդել</label>
            <input
              type="text"
              value={formData.car_model}
              onChange={(e) => handleFormChange('car_model', e.target.value)}
              className="form-input"
            />
          </div>

          <div className="form-group">
            <label>Ավտոմեքենայի համար</label>
            <input
              type="text"
              value={formData.car_number}
              onChange={(e) => handleFormChange('car_number', e.target.value)}
              className="form-input"
            />
          </div>

          <div className="form-group">
            <label>Ձիաուժ (HP)</label>
            <input
              type="text"
              value={formData.hp}
              onChange={(e) => handleFormChange('hp', e.target.value)}
              className="form-input"
            />
          </div>

          <div className="form-group">
            <label>Ժամանակահատված</label>
            <select
              value={formData.period}
              onChange={(e) => handleFormChange('period', e.target.value)}
              className="form-input"
            >
              <option value="">Ընտրել</option>
              <option value="long">Երկար</option>
              <option value="short">Կարճ</option>
            </select>
          </div>

          <div className="form-group">
            <label>Լրացուցիչ տեղեկություն</label>
            <input
              type="text"
              value={formData.info}
              onChange={(e) => handleFormChange('info', e.target.value)}
              className="form-input"
            />
          </div>

          <div className="form-group">
            <label>Վճար</label>
            <input
              type="number"
              step="0.01"
              value={formData.price}
              onChange={(e) => handleFormChange('price', e.target.value)}
              className="form-input"
            />
          </div>
           <div className="form-group">
            <label>Զեղչ(Գումար)</label>
            <input
              type="number"
              value={formData.discount}
              onChange={(e) => handleFormChange('discount', e.target.value)}
              className="form-input"
            />
          </div>
        </div>

        {message && (
          <div className={`form-message ${message.type}`}>
            {message.text}
          </div>
        )}

        <div className="form-buttons">
          <button 
            type="submit" 
            className="submit-btn"
            disabled={submitting}
          >
            {submitting ? 'Ընդամենը...' : 'Ավելացնել'}
          </button>
          <button 
            type="button" 
            className="reset-btn"
            onClick={() => {
              resetForm()
              onClose()
            }}
            disabled={submitting}
          >
            Չեղարկել
          </button>
        </div>
      </form>
    </div>
  )
}

export default InsurancePolicyForm
