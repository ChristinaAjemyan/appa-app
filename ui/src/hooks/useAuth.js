import { useState, useEffect } from 'react'

/**
 * Custom hook to get current user information from localStorage
 * and handle logout
 */
function useAuth() {
  const [user, setUser] = useState(null)
  const [token, setToken] = useState(null)
  const [isAuthenticated, setIsAuthenticated] = useState(false)

  useEffect(() => {
    const storedToken = localStorage.getItem('token')
    const storedUser = localStorage.getItem('user')

    if (storedToken && storedUser) {
      setToken(storedToken)
      setUser(JSON.parse(storedUser))
      setIsAuthenticated(true)
    } else {
      setIsAuthenticated(false)
    }
  }, [])

  const logout = () => {
    localStorage.removeItem('token')
    localStorage.removeItem('user')
    setToken(null)
    setUser(null)
    setIsAuthenticated(false)
  }

  const login = (token, userData) => {
    localStorage.setItem('token', token)
    localStorage.setItem('user', JSON.stringify(userData))
    setToken(token)
    setUser(userData)
    setIsAuthenticated(true)
  }

  const hasRole = (role) => {
    if (!isAuthenticated || !user) return false
    const rolesArray = Array.isArray(role) ? role : [role]
    return rolesArray.includes(user.role)
  }

  return {
    user,
    token,
    isAuthenticated,
    logout,
    login,
    hasRole
  }
}

export default useAuth
