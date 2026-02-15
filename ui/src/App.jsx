import { BrowserRouter as Router, Routes, Route } from 'react-router-dom'
import Layout from './components/Layout'
import Dashboard from './components/Dashboard'
import Agents from './components/Agents'
import InsurancePolicies from './components/InsurancePolicies'
import Settings from './components/Settings'
import AgentsPercentages from './components/AgentsPercentages'
import CompaniesPercentages from './components/CompaniesPercentages'
import Companies from './components/Companies'
import './App.css'

function App() {
  return (
    <Router>
      <Routes>
        <Route path="/" element={<Layout />}>
          <Route index element={<Dashboard />} />
          <Route path="/agents" element={<Agents />} />
          <Route path="/insurance-policies" element={<InsurancePolicies />} />
          <Route path="/settings" element={<Settings />}>
            <Route path="/settings/agents-percentages" index element={<AgentsPercentages />} />
            <Route path="/settings/companies-percentages" element={<CompaniesPercentages />} />
            <Route path="/settings/companies" element={<Companies />} />
          </Route>
        </Route>
      </Routes>
    </Router>
  )
}

export default App
