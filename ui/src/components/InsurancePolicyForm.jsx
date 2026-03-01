import { useState } from 'react'
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

  const handleFormChange = (field, value) => {
    setFormData(prev => ({ ...prev, [field]: value }))
  }

  const handleFormSubmit = async (e) => {
    e.preventDefault()
    try {
      setSubmitting(true)
      setMessage(null)

      // Prepare data for submission
      const submitData = {
        ...formData,
        price: formData.price ? parseFloat(formData.price) : null
      }

      await axios.post('/api/insurance-policies', submitData)
      
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
  }

  return (
    <div className="form-section">
      <h3>Նոր պայմանագիր ավելացնել</h3>
      <form onSubmit={handleFormSubmit} className="insurance-form">
        <div className="form-grid">
          <div className="form-group">
            <label>Ընկերություն *</label>
            <input
              type="text"
              value={formData.company}
              onChange={(e) => handleFormChange('company', e.target.value)}
              className="form-input"
              required
            />
          </div>

          <div className="form-group">
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
          </div>

          <div className="form-group">
            <label>Գործակալի անուն</label>
            <input
              type="text"
              value={formData.agent_name}
              onChange={(e) => handleFormChange('agent_name', e.target.value)}
              className="form-input"
            />
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
            <label>Տիրապետի անուն *</label>
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
              className="form-input"
            />
          </div>

          <div className="form-group">
            <label>Ավարտի ամսաթիվ</label>
            <input
              type="date"
              value={formData.end_date}
              onChange={(e) => handleFormChange('end_date', e.target.value)}
              className="form-input"
            />
          </div>

          <div className="form-group">
            <label>Շրջան</label>
            <input
              type="text"
              value={formData.region}
              onChange={(e) => handleFormChange('region', e.target.value)}
              className="form-input"
            />
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
            <label>Ձիաց ուժ (HP)</label>
            <input
              type="text"
              value={formData.hp}
              onChange={(e) => handleFormChange('hp', e.target.value)}
              className="form-input"
            />
          </div>

          <div className="form-group">
            <label>Ժամանակաշրջան</label>
            <input
              type="text"
              value={formData.period}
              onChange={(e) => handleFormChange('period', e.target.value)}
              className="form-input"
            />
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
            <label>Գինը</label>
            <input
              type="number"
              step="0.01"
              value={formData.price}
              onChange={(e) => handleFormChange('price', e.target.value)}
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
