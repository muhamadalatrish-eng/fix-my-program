import React from 'react'
import ReactDOM from 'react-dom/client'
import App from './App'
import './index.css'

// تفعيل الثيم الداكن افتراضياً على مستوى الجذور
document.documentElement.classList.add('dark');

ReactDOM.createRoot(document.getElementById('root')!).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
)