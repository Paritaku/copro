import { BrowserRouter, Routes, Route } from 'react-router-dom'
import './App.css'
import Layout from './components/components/Layout'
import GenererFichier from './pages/GenererFichier'

function App() {
  return (
    <BrowserRouter basename="">
      <Routes>
        <Route path='/' element={<Layout />}>
          <Route  index path='' element={<GenererFichier />} />
          <Route path='parametres' element={<div>Settings Page</div>} /> 
        </Route>
      </Routes>
    </BrowserRouter>  
  )
}

export default App
