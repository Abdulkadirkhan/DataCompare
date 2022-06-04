import React,{useState, useEffect} from 'react';
import { Link } from 'react-router-dom';
import { Button } from '../components/objects/Button';
import './Button.css';
import './Navbar.css';
import {useHistory} from 'react-router-dom';
import Dropdown from '../components/objects/Dropdown'


function Navbar() {
  const[click,setClick] = useState(false);
  const [button,setButton] = useState(true)
  const [dropdown, setDropdown] = useState(false);

  const handleClick = () => setClick(!click);
  const closeMobileMenu = () => setClick(false)

  const history = useHistory();
  const handleOnClick = () => history.push('/support');

  const onMouseEnter = () => {
    if (window.innerWidth < 960) {
      setDropdown(false);
    } else {
      setDropdown(true);
    }
  };

  const onMouseLeave = () => {
    if (window.innerWidth < 960) {
      setDropdown(false);
    } else {
      setDropdown(false);
    }
  };

  const showButton = () => {
    if(window.innerWidth <= 960) {
      setButton(false);
    } else {
      setButton(true);
    }
  };
  useEffect(()=> {
    showButton();
  },[]);

  window.addEventListener('resize', showButton)
    return (
      <>
      <nav className="navbar">
        <div className="navbar-container">
          <Link to="/" className="navbar-logo" onClick={closeMobileMenu}>
            @Capgemini <i className='i.fab.fa-typo3' />  
          </Link>
          <div className='menu-icon' onClick={handleClick}>
            <i className={ click ? 'fas fa-times' : 'fas fa-bars'} />
            </div>
            <ul className={click ? 'nav-menu active' : 'nav-menu'}>
              <li className='nav-item'>
                <Link to='/' className='nav-links' onClick={closeMobileMenu}>
                  Home
                </Link>
              </li>
              <li 
              className='nav-item'
              onMouseEnter={onMouseEnter}
              onMouseLeave={onMouseLeave}
              >
                <Link to='/services' className='nav-links' onClick={closeMobileMenu}>
                  Sevices <i className='fas fa-caret-down' />
                </Link>
                {dropdown && <Dropdown />}
              </li>
              <li className='nav-item'>
                <Link to='/docs' className='nav-links' onClick={closeMobileMenu}>
                  Docs
                </Link>
              </li>
              <li>
              <Link to='/support' className='nav-links-mobile' onClick={closeMobileMenu}
              >
                Support
              </Link>
            </li>

            </ul>
            
            {button && <Button buttonStyle='btn--outline' onClick={handleOnClick} >Support</Button>}
            


        </div>
      </nav>
      </>
    );
}

export default Navbar
