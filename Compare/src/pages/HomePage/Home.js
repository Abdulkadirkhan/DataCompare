import React from 'react';
import '../../App.css';
import HomeSection from '../../components/sections/HomeSection';
import { homeObjOne} from './Data';
import { InfoSection,Service} from '../../components';
import Docs from '../Documents/docs';
import Support from '../Support/support';

function Home() {
  return (
    <>
      
      <InfoSection {...homeObjOne} />
      <Service />
      <Docs/>
      <Support/>
    </>
  );
}

export default Home;
