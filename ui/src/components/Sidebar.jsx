import { Link, useLocation } from 'react-router-dom'
import { useState } from 'react'
import './Sidebar.css'

function Sidebar() {
  const location = useLocation()
  const [expandedItems, setExpandedItems] = useState({})

  const menuItems = [
    { path: '/', label: 'Հաշվարկ', icon: '📊' },
    { path: '/agents', label: 'Գործակալներ', icon: '💰' },
    { path: '/insurance-policies', label: 'Պայմանագրեր', icon: '📋' },
    { 
      path: '/settings', 
      label: 'Կարգավորումներ', 
      icon: '⚙️',
      children: [
        { path: '/settings/agents-percentages', label: 'Գործակալների տոկոսներ', icon: '📈' },
        { path: '/settings/companies', label: 'Ընկերություններ', icon: '🏢' }
      ]
    },
  ]

  const toggleExpand = (path) => {
    setExpandedItems(prev => ({
      ...prev,
      [path]: !prev[path]
    }))
  }

  const isActive = (path) => location.pathname === path || location.pathname.startsWith(path + '/')

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
    </aside>
  )
}

export default Sidebar
