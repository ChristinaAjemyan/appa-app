import { Outlet, useLocation, Link } from 'react-router-dom'
import './Settings.css'

function Settings() {
  const location = useLocation()

  const submenu = [
    { path: '/settings/agents-percentages', label: 'Գործակալների տոկոսներ' },
    { path: '/settings/companies-percentages', label: 'Ընկերությունների տոկոսներ' },
    { path: '/settings/companies', label: 'Ընկերություններ' },
  ]

  return (
    <div className="settings">
      <div className="settings-header">
        <h2>Կարգավորումներ</h2>
      </div>

      <div className="submenu-tabs">
        {submenu.map((item) => (
          <Link
            key={item.path}
            to={item.path}
            className={`submenu-tab ${location.pathname === item.path ? 'active' : ''}`}
          >
            {item.label}
          </Link>
        ))}
      </div>

      <div className="submenu-content">
        <Outlet />
      </div>
    </div>
  )
}

export default Settings
