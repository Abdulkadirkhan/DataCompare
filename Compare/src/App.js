import React from 'react';
//import Navbar from './components/Navbar';
import { Navbar, Footer } from './components';
import Services from './pages/Services/Services';
import Excelcompare from './pages/Services/excelcompare';
import Upload from './pages/Services/uploadMaster';
import SingleExcelCompare from './pages/Services/fileexcelcompare';
import './App.css';
import Home from './pages/HomePage/Home';
import { BrowserRouter as Router, Switch, Route } from 'react-router-dom';
import Docs from './pages/Documents/docs';
import Support from './pages/Support/support';
import GlobalStyle from './globalStyles';

function App() {
  return (
    <>
      <Router>
        <GlobalStyle />
        <Navbar />
        <div style={{minHeight: '75vh'}}>
        <Switch>
          <Route path='/' exact component={Home} />
          <Route path='/services' component={Services} />
          <Route path='/excelcompare' component={Excelcompare} />
          <Route path='/upload' component={Upload} />
          <Route path='/fileexcelcompare' component={SingleExcelCompare} />
          <Route path='/docs' component={Docs} />
          <Route path='/support' component={Support} />
        </Switch>
        </div>
        <Footer />
      </Router>
    </>
  );
}

export default App;
