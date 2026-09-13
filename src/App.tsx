import { useState } from 'react';
import { LoginScreen } from './components/login-screen';
import { AppShell } from './components/app-shell';

export default function App() {
  const [isAuthenticated, setIsAuthenticated] = useState(false);

  if (!isAuthenticated) {
    return <LoginScreen onLogin={() => setIsAuthenticated(true)} />;
  }

  return <AppShell />;
}