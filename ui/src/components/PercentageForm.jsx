import { useState, useEffect } from 'react'
import './PercentageForm.css'

function PercentageForm({ item, onSave, onCancel, isLoading }) {
  const [formData, setFormData] = useState({
    company: '',
    product_type: '',
    agent_code_in: '',
    agent_code_not: '',
    region_in: '',
    region_not: '',
    bm_min: '',
    bm_max: '',
    bm_exact: '',
    brand_in: '',
    hp_min: '',
    hp_max: '',
    period: '',
    percentage: '',
  })

  useEffect(() => {
    if (item) {
      setFormData(item)
    }
  }, [item])

  const handleChange = (e) => {
    const { name, value } = e.target
    setFormData(prev => ({
      ...prev,
      [name]: value
    }))
  }

  const handleSubmit = (e) => {
    e.preventDefault()
    if (!formData.company || formData.percentage === '') {
      alert('Company and Percentage are required')
      return
    }
    onSave(formData)
  }

  return (
    <form className="percentage-form" onSubmit={handleSubmit}>
      <div className="form-group">
        <label>Ընկերություն *</label>
        <input
          type="text"
          name="company"
          value={formData.company}
          onChange={handleChange}
          placeholder="Company name"
          required
        />
      </div>

      <div className="form-group">
        <label>Պրոդուկտի Տեսակ</label>
        <input
          type="text"
          name="product_type"
          value={formData.product_type}
          onChange={handleChange}
          placeholder="Product type(APPA,APKA)"
        />
      </div>

      <div className="form-row">
        <div className="form-group">
          <label>Գործակալ Կոդ (In)</label>
          <input
            type="text"
            name="agent_code_in"
            value={formData.agent_code_in}
            onChange={handleChange}
            placeholder="e.g. 768%"
          />
        </div>
        <div className="form-group">
          <label>Գործակալ Կոդ (Not)</label>
          <input
            type="text"
            name="agent_code_not"
            value={formData.agent_code_not}
            onChange={handleChange}
            placeholder="Exclude codes"
          />
        </div>
      </div>

      <div className="form-row">
        <div className="form-group">
          <label>Շրջան (In)</label>
          <input
            type="text"
            name="region_in"
            value={formData.region_in}
            onChange={handleChange}
            placeholder="e.g. Երևան"
          />
        </div>
        <div className="form-group">
          <label>Շրջան (Not)</label>
          <input
            type="text"
            name="region_not"
            value={formData.region_not}
            onChange={handleChange}
            placeholder="Exclude regions"
          />
        </div>
      </div>

      <div className="form-row">
        <div className="form-group">
          <label>BM Min</label>
          <input
            type="number"
            name="bm_min"
            value={formData.bm_min}
            onChange={handleChange}
            placeholder="Min"
          />
        </div>
        <div className="form-group">
          <label>BM Max</label>
          <input
            type="number"
            name="bm_max"
            value={formData.bm_max}
            onChange={handleChange}
            placeholder="Max"
          />
        </div>
        <div className="form-group">
          <label>BM Exact</label>
          <input
            type="number"
            name="bm_exact"
            value={formData.bm_exact}
            onChange={handleChange}
            placeholder="Exact"
          />
        </div>
      </div>

      <div className="form-row">
        <div className="form-group">
          <label>Brand In</label>
          <input
            type="text"
            name="brand_in"
            value={formData.brand_in}
            onChange={handleChange}
            placeholder="e.g. BMW, Honda"
          />
        </div>
        <div className="form-group">
          <label>Period</label>
          <input
            type="text"
            name="period"
            value={formData.period}
            onChange={handleChange}
            placeholder="e.g. Երկար, Կարճ"
          />
        </div>
      </div>

      <div className="form-row">
        <div className="form-group">
          <label>HP Min</label>
          <input
            type="number"
            name="hp_min"
            value={formData.hp_min}
            onChange={handleChange}
            placeholder="Min HP"
          />
        </div>
        <div className="form-group">
          <label>HP Max</label>
          <input
            type="number"
            name="hp_max"
            value={formData.hp_max}
            onChange={handleChange}
            placeholder="Max HP"
          />
        </div>
      </div>

      <div className="form-group">
        <label>Տոկոս *</label>
        <input
          type="number"
          name="percentage"
          value={formData.percentage}
          onChange={handleChange}
          placeholder="Percentage"
          step="0.01"
          required
        />
      </div>

      <div className="form-actions">
        <button type="button" className="btn-cancel" onClick={onCancel} disabled={isLoading}>
          Չեղարկել
        </button>
        <button type="submit" className="btn-save" disabled={isLoading}>
          {isLoading ? 'Պահպանում...' : 'Պահպանել'}
        </button>
      </div>
    </form>
  )
}

export default PercentageForm
