import React from 'react';

import {
  FooterContainer,
  SocialMedia,
  SocialMediaWrap,
  WebsiteRights
} from './Footer.elements';

function Footer() {
  return (
    <FooterContainer>
      <SocialMedia>
        <SocialMediaWrap>
          
          <WebsiteRights sx={{ textAlign : 'right'}}>CAPGEMINI © 2022</WebsiteRights>
          
        </SocialMediaWrap>
      </SocialMedia>
    </FooterContainer>
  );
}

export default Footer;
