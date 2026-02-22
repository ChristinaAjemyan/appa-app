import { useState, useEffect } from 'react'
import './PercentageForm.css'

function AgentForm({ agent, onSave, onCancel, isLoading }) {
  const [formData, setFormData] = useState({
    code: '',
    name: '',
    phone_number: '',
    NAIRI: '',
    REGO: '',
    ARMENIA: '',
    SIL: '',
    INGO: '',
    LIGA: '',
  })

  useEffect(() => {
    if (agent) {
      setFormData(agent)
    }
  }, [agent])

  const handleChange = (e) => {
    const { name, value } = e.target
    setFormData(prev => ({
      ...prev,
      [name]: value
    }))
  }

  const handleSubmit = (e) => {
    e.preventDefault()
    if (!formData.code || !formData.name) {
      alert('Կոդ և անուն պարտադիր են')
      return
    }
    onSave(formData)
  }

  return (
    <form className="percentage-form" onSubmit={handleSubmit}>
      <div className="form-row">
        <div className="form-group">
          <label>Կոդ *</label>
          <input
            type="text"
            name="code"
            value={formData.code}
            onChange={handleChange}
            placeholder="e.g. 768-01"
            required
          />
        </div>
        <div className="form-group">
          <label>Անուն *</label>
          <input
            type="text"
            name="name"
            value={formData.name}
            onChange={handleChange}
            placeholder="Agent name"
            required
          />
        </div>
      </div>

      <div className="form-group">
        <label>Հեռ․ համար</label>
        <input
          type="text"
          name="phone_number"
          value={formData.phone_number}
          onChange={handleChange}
          placeholder="Phone number"
        />
      </div>

      <div className="form-row">
        <div className="form-group">
          <label>NAIRI</label>
          <input
            type="text"
            name="NAIRI"
            value={formData.NAIRI}
            onChange={handleChange}
            placeholder="NAIRI code"
          />
        </div>
        <div className="form-group">
          <label>REGO</label>
          <input
            type="text"
            name="REGO"
            value={formData.REGO}
            onChange={handleChange}
            placeholder="REGO code"
          />
        </div>
      </div>

      <div className="form-row">
        <div className="form-group">
          <label>ARMENIA</label>
          <input
            type="text"
            name="ARMENIA"
            value={formData.ARMENIA}
            onChange={handleChange}
            placeholder="ARMENIA code"
          />
        </div>
        <div className="form-group">
          <label>SIL</label>
          <input
            type="text"
            name="SIL"
            value={formData.SIL}
            onChange={handleChange}
            placeholder="SIL code"
          />
        </div>
      </div>

      <div className="form-row">
        <div className="form-group">
          <label>INGO</label>
          <input
            type="text"
            name="INGO"
            value={formData.INGO}
            onChange={handleChange}
            placeholder="INGO code"
          />
        </div>
        <div className="form-group">
          <label>LIGA</label>
          <input
            type="text"
            name="LIGA"
            value={formData.LIGA}
            onChange={handleChange}
            placeholder="LIGA code"
          />
        </div>
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

export default AgentForm
