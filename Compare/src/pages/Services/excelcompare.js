import React from 'react';
import { Button } from '../../globalStyles';
//import { AiFillThunderbolt } from 'react-icons/ai';
import { GiCrystalBars } from 'react-icons/gi';
import { GiCutDiamond, GiRock,GiAerodynamicHarpoon } from 'react-icons/gi';
//import { GiFloatingCrystal } from 'react-icons/gi';
import { IconContext } from 'react-icons/lib';
import {
  ServiceSection,
  ServiceWrapper,
  ServiceHeading,
  ServiceContainer,
  ServiceCard,
  ServiceCardInfo,
  ServiceCardIcon,
  ServiceCardPlan,
  ServiceCardCost,
  ServiceCardLength,
  ServiceCardFeatures,
  ServiceCardFeature
} from './excelcompare.elements';

function ExcelCompare() {
  return (
    <IconContext.Provider value={{ color: '#a9b3c1', size: 64 }}>
      <ServiceSection>
        <ServiceWrapper>
          <ServiceHeading>Choose Type</ServiceHeading>
          <ServiceContainer>
            <ServiceCard to='/fileexcelcompare'>
              <ServiceCardInfo>
                <ServiceCardIcon>
                  
                </ServiceCardIcon>
                <ServiceCardPlan>Single File Comparison</ServiceCardPlan>
                
                <Button primary>START</Button>
              </ServiceCardInfo>
            </ServiceCard>
            <ServiceCard to='/upload'>
              <ServiceCardInfo>
                <ServiceCardIcon>
             
                </ServiceCardIcon>
                <ServiceCardPlan>Multiple Files Comparsion</ServiceCardPlan>
                <Button primary>START</Button>
              </ServiceCardInfo>
            </ServiceCard>
            
          </ServiceContainer>
        </ServiceWrapper>
      </ServiceSection>
    </IconContext.Provider>
  );
}
export default ExcelCompare;
