import { Navigate } from 'react-router-dom'

/**
 * ProtectedRoute component that checks if user is authenticated
 * and optionally checks if user has required role
 * @param {React.ReactNode} children - Component to render if authorized
 * @param {string|string[]} allowedRoles - Role(s) allowed to access this route
 */
function ProtectedRoute({ children, allowedRoles }) {
  const token = localStorage.getItem('token')
  const user = JSON.parse(localStorage.getItem('user') || '{}')

  // Redirect to login if no token
  if (!token) {
    return <Navigate to="/login" replace />
  }

  // Check role-based access if allowedRoles is specified
  if (allowedRoles) {
    const rolesArray = Array.isArray(allowedRoles) ? allowedRoles : [allowedRoles]
    const hasRole = rolesArray.includes(user.role)
    
    if (!hasRole) {
      // Redirect to insurance policies (default accessible page) or home
      return <Navigate to="/insurance-policies" replace />
    }
  }

  return children
}

export default ProtectedRoute
