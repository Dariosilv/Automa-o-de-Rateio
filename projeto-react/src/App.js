import React from 'react';
import './App.css';
import ExcelManager from './components/ExcelManager';

function App() {
  return (
    <div className="App">
      <header className="App-header">
        <h3>Automação de Rateio de E-mails</h3>
        <ExcelManager />
      </header>
    </div>
  );
}

export default App;
