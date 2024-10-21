/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
//loki auth token from session storage,
// eyJhbGciOiJSU0EtT0FFUCIsImVuYyI6IkExMjhDQkMtSFMyNTYiLCJ4NXQiOiJtRXpHeF80OE9MOWcxdXh1QkFnRXVOdGJqcVEiLCJ6aXAiOiJERUYifQ.J0GAJX7SnDfdsoXMboB4Vm8G4jTUs3aBPWCycFe4AlnK-LXxRdKmKDLswEDAetH_FK-M64kQzDmfbdqGR-BK2gFf9jsrA2T2QMJfOuYVH2mEKotr7MRnkYTRyPXpQ4xfyvjT3J2W9fNpF0OSENHcU872_fNXxqtdOXpqQuM7tbDPmrdkdkDNWNFiDxFJaPYRJBY-ITXnbCUv7Q6V9pVS5XGUKvKKu5iW7Zy3IGRDpXaQz8tjAWQPhkyNmXRgibrzAqHa_C4PdlKuQg-RuJM69I-u1rcYXxvUDWTpsNme5fRaPGbg73eli_h6uux4bFsvBYVvnizA3Sjapxdc8x1oTg.IzNlswzsjOZ9M69rh3IpFA.xpgqFsFOLIpMh2brsqa4MIJUr0uR52QYHROeQSufglYeyWqE7aVUpqDDu6dV8J4RE1Elvc_5wXkMYCcoLJF_vIaZ8MW8EYMipdR8Cj4zSpSvLWa29U0NcT5_pEnw3qitzBGQm-xm0eUwmZJMoykOzpmaRHCrXtv-bCbuRqI9ra9nzMtnsQIfyEFTSP8QfSX4xvEfezQZoZHeuQWVpykygp_HMaf7xLnvQwmoTfn9nyfDO6S_E9qEu_EtsBjUo2JlfyQfL2ErlZ0raEu8NbVd5wCIldFMl2zkGZ4P6JUW3fLmKqnrn_SgoW7X47wPJ5V-SrMNlaAi3ljQvvnfeSSqU9374V87GT2xV7mRYud-TaTMneqq0DdF3ScGlHnbfsBWPxqg3XAZSyInNDGUkbbhuRLZc3akky4oGnM-LNDA4YG476Mk6E7fod8YEMIfy8D8GeXg6NO3agY2n46CBG55tzdoilqFAnwK40ml4PvuX5pvBfTA0_56CiAtdVCbm5TTGuXr2X3PDfXNI1oHEnC3nfq5la8CUc9E64veNiM66Rebefy2m3KSJQ2yHI8f7VHHxRtV7y_skEXxLmsP87JJLGQ_vnb366hJjfa7iMaOWo0l7hLfv4u4fq_JzkZsL96AKPbwnDtY6W8eBjmdHLO3rs72YzPDmZNMm7JzwRtE3S2eo2QLEqnQKHCCRPkORSr0r3i7fOw_4A5CUU35tMyBtHgVOzkRMQ-Vcjf7HIv5-7sJMF2-d1LiadFMJVKKtmwd3PtqwtR1VPMC6DUpPKhFoHyHdf4t7EQIH5litSfCsZG7UfqE_R4qw5X_kIFM8_kO13CcxMY-Oxlw5kCQT7CWrpsuWIO6I9kRYImrYohmrpaAnWupni7DuoTfb4vXRT69WHhjwChS-GLa27qAvg6BvZl4kChO6O5STiDRny7f9ebFE41Dr8jFL7EOtL0tFIJy5azlOtdH-gtU-uiIpejggBULinh2qsP8aM6ywu6ozbIsRx4D77FFLRsyLwc5hjWYn8S-M-_9H6_b1Z92oSai0HIkUEWn2k72UIH9kTwxK88tYzOWhwC9yrcHaBs33sleQcCInhytk3VUd0UmWDwz659y-mjVq5SbM4dChkjTfI0TIagxulPU8LedKwJ2uLHBKOgZx9N2fZHpiBje4S0iYrFdoC4y_61tJ-0T209rIeLEM3L9FTQsVkxlGqNnPNwe2ocVaLVehn4DSJZ_qmWFe4VrkBc6oGwnD-5GaqnlBSCVII3ex94mHpSBXHpmZRZxI1gJx7xjECAlEu-8LQ45vdqWZ5osSp3a1GN31xf2dWguwPMIUuTLAyEViBjxsMX4QDfyHTgQ7tGMAZDp3ELH49FKjPCL2bO1unawOf_h42fjeImqq2xAZVP_fQM6CuWLQ9CMD1jWyGnNOWj2foukVnESDsfQ7h_NKztFludRyV-eN8E4sXBIxPyzgOi-_XTx7pqE0cbiuXSsGIl_i2VJx2hjrINPPuf0RlRc2IzHhu9-WfVTwbC5IKLPerfNQz6qjeKDvq8qjNJFbj51jSWnVkeQEqulrf7iF7CJ_d3xAGFfZzWk1zj0_VAxgWdxJCSSTOMPr4FdXAfICAJ3qBEklZFSkJXoEpc8KLhkOBBtbAD3HIGsKS_JE9-bOMfvrDHo6xMpeNPB3WrHI5FtPjhIGw.axRwkUzM8ZIHrj0u_xpBKg
import { getUserProfile } from "../helpers/sso-helper";
import { filterUserProfileInfo } from "./../helpers/documentHelper";
import {getOrganisationData} from '../helpers/lists-helper'
Office.onReady((info) => {
  
});


function loadData() {
  getOrganisationData()
}

export async function run() {
  //getUserProfile(writeDataToOfficeDocument);
}

function writeDataToOfficeDocument(result) {
  /*let data = [];
  let userProfileInfo = filterUserProfileInfo(result);

  for (let i = 0; i < userProfileInfo.length; i++) {
    if (userProfileInfo[i] !== null) {
      data.push(userProfileInfo[i]);
    }
  }

  let userInfo = "";
  for (let i = 0; i < data.length; i++) {
    userInfo += data[i] + "\n";
  }

  document.getElementById("main").style.display = "none";
  Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
  */
}
