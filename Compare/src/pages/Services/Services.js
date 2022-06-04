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
} from './Service.elements';

function Services() {
  return (
    <IconContext.Provider value={{ color: '#a9b3c1', size: 64 }}>
      <ServiceSection>
        <ServiceWrapper>
          <ServiceHeading>Select Service To Start</ServiceHeading>
          <ServiceContainer>
            <ServiceCard to='/excelcompare'>
              <ServiceCardInfo>
                <ServiceCardIcon>
                  <GiRock />
                </ServiceCardIcon>
                <ServiceCardPlan>EXCEL Comparsion</ServiceCardPlan>
                <ServiceCardLength>General Solution Features & Options:</ServiceCardLength>
                <ServiceCardFeatures>
                  <ServiceCardFeature>1-To-1 Comparison Mode</ServiceCardFeature>
                  <ServiceCardFeature>Compare folders with master.xls</ServiceCardFeature>
                  <ServiceCardFeature>Multiple Filtering Options</ServiceCardFeature>
                  <ServiceCardFeature>Detailed and summary reports</ServiceCardFeature>
                </ServiceCardFeatures>
                <Button primary>BEGIN</Button>
              </ServiceCardInfo>
            </ServiceCard>
            <ServiceCard to='/pdfcompare'>
              <ServiceCardInfo>
                <ServiceCardIcon>
                <GiCrystalBars />
                </ServiceCardIcon>
                <ServiceCardPlan>PDF Comparsion</ServiceCardPlan>
                <ServiceCardLength>General Solution Features & Options:</ServiceCardLength>
                <ServiceCardFeatures>
                  <ServiceCardFeature>1-To-1 Comparison Mode</ServiceCardFeature>
                  <ServiceCardFeature>Compare folders with master.xls</ServiceCardFeature>
                  <ServiceCardFeature>Multiple Filtering Options</ServiceCardFeature>
                  <ServiceCardFeature>Detailed and summary reports</ServiceCardFeature>
                </ServiceCardFeatures>
                <Button primary>BEGIN</Button>
              </ServiceCardInfo>
            </ServiceCard>
            <ServiceCard to='/imagecompare'>
              <ServiceCardInfo>
                <ServiceCardIcon>
                <GiCutDiamond />
                </ServiceCardIcon>
                <ServiceCardPlan>IMAGE Comparsion</ServiceCardPlan>
                <ServiceCardLength>General Solution Features & Options:</ServiceCardLength>
                <ServiceCardFeatures>
                  <ServiceCardFeature>1-To-1 Comparison Mode</ServiceCardFeature>
                  <ServiceCardFeature>Compare folders with master.xls</ServiceCardFeature>
                  <ServiceCardFeature>Multiple Filtering Options</ServiceCardFeature>
                  <ServiceCardFeature>Detailed and summary reports</ServiceCardFeature>
                </ServiceCardFeatures>
                <Button primary>BEGIN</Button>
              </ServiceCardInfo>
            </ServiceCard>
            <ServiceCard to='/textcompare'>
              <ServiceCardInfo>
                <ServiceCardIcon>
                <GiAerodynamicHarpoon/>
                </ServiceCardIcon>
                <ServiceCardPlan>TEXT Comparsion</ServiceCardPlan>
                <ServiceCardLength>General Solution Features & Options:</ServiceCardLength>
                <ServiceCardFeatures>
                  <ServiceCardFeature>1-To-1 Comparison Mode</ServiceCardFeature>
                  <ServiceCardFeature>Compare folders with master.xls</ServiceCardFeature>
                  <ServiceCardFeature>Multiple Filtering Options</ServiceCardFeature>
                  <ServiceCardFeature>Detailed and summary reports</ServiceCardFeature>
                </ServiceCardFeatures>
                <Button primary>BEGIN</Button>
              </ServiceCardInfo>
            </ServiceCard>
          </ServiceContainer>
        </ServiceWrapper>
      </ServiceSection>
    </IconContext.Provider>
  );
}
export default Services;
