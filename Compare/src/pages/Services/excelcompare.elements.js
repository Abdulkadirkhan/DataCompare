import { Link } from 'react-router-dom';
import styled from 'styled-components';

export const ServiceSection = styled.div`
  padding: 100px 0 160px;
  display: flex;
  flex-direction: column;
  justify-content: center;
  background: #4b59f7;
`;

export const ServiceWrapper = styled.div`
  display: flex;
  flex-direction: column;
  align-items: center;
  margin: 0 auto;

  @media screen and (max-width: 960px) {
    margin: 0 30px;
    display: flex;
    flex-direction: column;
    align-items: center;
  }
`;

export const ServiceHeading = styled.h1`
  color: #fff;
  font-size: 48px;
  margin-bottom: 24px;
`;

export const ServiceContainer = styled.div`
  display: flex;
  justify-content: center;
  align-items: center;

  @media screen and (max-width: 960px) {
    display: flex;
    flex-direction: column;
    justify-content: center;
    align-items: center;
    width: 100%;
  }
`;

export const ServiceCard = styled(Link)`
  background: whitesmoke;
  box-shadow: 0 6px 20px rgba(56, 125, 255, 0.2);
  width: 400px;
  height: 300px;
  text-decoration: none;
  border-radius: 4px;

  &:nth-child(2) {
    margin: 24px;
  }

  &:hover {
    transform: scale(1.06);
    transition: all 0.3s ease-out;
    color: #1c2237;
  }

  @media screen and (max-width: 960px) {
    width: 90%;

    &:hover {
      transform: none;
    }
  }
`;

export const ServiceCardInfo = styled.div`
  display: flex;
  flex-direction: column;
  height: 500px;
  padding: 24px;
  align-items: center;
  color: #fff;
`;

export const ServiceCardIcon = styled.div`
  margin: 24px 0;
`;

export const ServiceCardPlan = styled.h3`
  margin-bottom: 20px;
  font-size: 20px;
  font-family : 'Courier New', Courier, monospace;
  color: grey;
  
`;

export const ServiceCardCost = styled.h4`
  font-size: 40px;
`;

export const ServiceCardLength = styled.p`
  font-size: 12px;
  margin-bottom: 5px;
  padding: 1px;
 
`;

export const ServiceCardFeatures = styled.ul`
  margin: 16px 0 32px;
  list-style: none;
  display: flex;
  flex-direction: column;
  align-items: center;
  color: #a9b3c1;
`;

export const ServiceCardFeature = styled.li`
  margin-bottom: 10px;
`;
