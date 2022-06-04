import React from 'react';
import '../../App.css';
import { Button } from '../objects/Button';
import './HomeSection.css';
import { Link } from 'react-router-dom';

function HomeSection() {
  return (
    <div className='hero-container'>
     
      <h1>Compare</h1>
      <p>What are you waiting for?</p>
      <div className='hero-btns'>
      <Link to='/services'>
        <Button
          className='btns'
          buttonStyle='btn--outline'
          buttonSize='btn--large'
        >
          GET STARTED
        </Button>
       </Link>
      </div>
    </div>
  );
}

export default HomeSection;
