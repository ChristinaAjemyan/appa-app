import { Link, useLocation, useNavigate } from 'react-router-dom'
import { useState } from 'react'
import './Sidebar.css'

function Sidebar() {
  const location = useLocation()
  const navigate = useNavigate()
  const [expandedItems, setExpandedItems] = useState({})
  const user = JSON.parse(localStorage.getItem('user') || '{}')

  // Define all menu items with role restrictions
  const allMenuItems = [
    { path: '/', label: 'Հաշվարկ', icon: '📊', roles: ['admin'] },
    { path: '/agents', label: 'Գործակալներ', icon: '💰', roles: ['admin'] },
    { path: '/insurance-policies', label: 'Պայմանագրեր', icon: '📋', roles: ['admin', 'employee'] },
    { 
      path: '/settings', 
      label: 'Կարգավորումներ', 
      icon: '⚙️',
      roles: ['admin'],
      children: [
        { path: '/settings/agents-percentages', label: 'Գործակալների տոկոսներ', icon: '📈' },
        { path: '/settings/companies', label: 'Ընկերություններ', icon: '🏢' }
      ]
    },
  ]

  // Filter menu items based on user role
  const menuItems = allMenuItems.filter(item => 
    item.roles.includes(user.role || 'employee')
  )

  const toggleExpand = (path) => {
    setExpandedItems(prev => ({
      ...prev,
      [path]: !prev[path]
    }))
  }

  const isActive = (path) => location.pathname === path || location.pathname.startsWith(path + '/')

  const handleLogout = () => {
    localStorage.removeItem('token')
    localStorage.removeItem('user')
    navigate('/login')
  }

  return (
    <aside className="sidebar">
      <div className="sidebar-header">
        <h1>Ապպա հաշվարկ </h1>
      </div>
      <nav className="sidebar-nav">
        <ul>
          {menuItems.map((item) => (
            <li key={item.path}>
              <div className="nav-item-wrapper">
                <Link
                  to={item.path}
                  className={`nav-link ${isActive(item.path) ? 'active' : ''}`}
                >
                  <span className="icon">{item.icon}</span>
                  <span className="label">{item.label}</span>
                </Link>
                {item.children && (
                  <button
                    className="expand-btn"
                    onClick={() => toggleExpand(item.path)}
                  >
                    {expandedItems[item.path] ? '▼' : '▶'}
                  </button>
                )}
              </div>
              {item.children && expandedItems[item.path] && (
                <ul className="submenu">
                  {item.children.map((child) => (
                    <li key={child.path}>
                      <Link
                        to={child.path}
                        className={`nav-link submenu-link ${location.pathname === child.path ? 'active' : ''}`}
                      >
                        <span className="icon">{child.icon}</span>
                        <span className="label">{child.label}</span>
                      </Link>
                    </li>
                  ))}
                </ul>
              )}
            </li>
          ))}
        </ul>
      </nav>
      <div className="sidebar-footer">
        <div className="user-info">
          <div className="user-avatar">{user.name?.charAt(0) || 'U'}</div>
          <div className="user-details">
            <p className="user-name">{user.name || 'User'}</p>
            <p className="user-role">{user.role || 'employee'}</p>
          </div>
        </div>
        <button className="logout-button" onClick={handleLogout}>
          Դուրս գալ
        </button>
      </div>
    </aside>
  )
}

export default Sidebar
