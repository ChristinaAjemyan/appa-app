import { BrowserRouter as Router, Routes, Route } from 'react-router-dom'
import Layout from './components/Layout'
import Dashboard from './pages/Dashboard/Dashboard'
import Agents from './pages/Agents/Agents'
import InsurancePolicies from './pages/InsurancePolicies/InsurancePolicies'
import Settings from './pages/Settings/Settings'
import AgentsPercentages from './pages/Settings/AgentsPercentages/AgentsPercentages'
import CompaniesPercentages from './pages/Settings/CompaniesPercentages/CompaniesPercentages'
import Companies from './pages/Settings/Companies/Companies'
import Login from './pages/Login/Login'
import Register from './pages/Register/Register'
import ProtectedRoute from './components/ProtectedRoute'
import './App.css'

function App() {
  return (
    <Router>
      <Routes>
        <Route path="/login" element={<Login />} />
        <Route path="/register" element={<Register />} />
        <Route
          path="/"
          element={
            <ProtectedRoute allowedRoles="admin">
              <Layout />
            </ProtectedRoute>
          }
        >
          <Route index element={<Dashboard />} />
          <Route path="/agents" element={<Agents />} />
          <Route path="/settings" element={<Settings />}>
            <Route path="/settings/agents-percentages" index element={<AgentsPercentages />} />
            <Route path="/settings/companies-percentages" element={<CompaniesPercentages />} />
            <Route path="/settings/companies" element={<Companies />} />
          </Route>
        </Route>
        <Route
          path="/insurance-policies"
          element={
            <ProtectedRoute allowedRoles={['admin', 'employee']}>
              <Layout />
            </ProtectedRoute>
          }
        >
          <Route index element={<InsurancePolicies />} />
        </Route>
      </Routes>
    </Router>
  )
}

export default App
