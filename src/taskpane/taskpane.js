/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
//loki auth token from session storage,
// eyJhbGciOiJSU0EtT0FFUCIsImVuYyI6IkExMjhDQkMtSFMyNTYiLCJ4NXQiOiJtRXpHeF80OE9MOWcxdXh1QkFnRXVOdGJqcVEiLCJ6aXAiOiJERUYifQ.J0GAJX7SnDfdsoXMboB4Vm8G4jTUs3aBPWCycFe4AlnK-LXxRdKmKDLswEDAetH_FK-M64kQzDmfbdqGR-BK2gFf9jsrA2T2QMJfOuYVH2mEKotr7MRnkYTRyPXpQ4xfyvjT3J2W9fNpF0OSENHcU872_fNXxqtdOXpqQuM7tbDPmrdkdkDNWNFiDxFJaPYRJBY-ITXnbCUv7Q6V9pVS5XGUKvKKu5iW7Zy3IGRDpXaQz8tjAWQPhkyNmXRgibrzAqHa_C4PdlKuQg-RuJM69I-u1rcYXxvUDWTpsNme5fRaPGbg73eli_h6uux4bFsvBYVvnizA3Sjapxdc8x1oTg.IzNlswzsjOZ9M69rh3IpFA.xpgqFsFOLIpMh2brsqa4MIJUr0uR52QYHROeQSufglYeyWqE7aVUpqDDu6dV8J4RE1Elvc_5wXkMYCcoLJF_vIaZ8MW8EYMipdR8Cj4zSpSvLWa29U0NcT5_pEnw3qitzBGQm-xm0eUwmZJMoykOzpmaRHCrXtv-bCbuRqI9ra9nzMtnsQIfyEFTSP8QfSX4xvEfezQZoZHeuQWVpykygp_HMaf7xLnvQwmoTfn9nyfDO6S_E9qEu_EtsBjUo2JlfyQfL2ErlZ0raEu8NbVd5wCIldFMl2zkGZ4P6JUW3fLmKqnrn_SgoW7X47wPJ5V-SrMNlaAi3ljQvvnfeSSqU9374V87GT2xV7mRYud-TaTMneqq0DdF3ScGlHnbfsBWPxqg3XAZSyInNDGUkbbhuRLZc3akky4oGnM-LNDA4YG476Mk6E7fod8YEMIfy8D8GeXg6NO3agY2n46CBG55tzdoilqFAnwK40ml4PvuX5pvBfTA0_56CiAtdVCbm5TTGuXr2X3PDfXNI1oHEnC3nfq5la8CUc9E64veNiM66Rebefy2m3KSJQ2yHI8f7VHHxRtV7y_skEXxLmsP87JJLGQ_vnb366hJjfa7iMaOWo0l7hLfv4u4fq_JzkZsL96AKPbwnDtY6W8eBjmdHLO3rs72YzPDmZNMm7JzwRtE3S2eo2QLEqnQKHCCRPkORSr0r3i7fOw_4A5CUU35tMyBtHgVOzkRMQ-Vcjf7HIv5-7sJMF2-d1LiadFMJVKKtmwd3PtqwtR1VPMC6DUpPKhFoHyHdf4t7EQIH5litSfCsZG7UfqE_R4qw5X_kIFM8_kO13CcxMY-Oxlw5kCQT7CWrpsuWIO6I9kRYImrYohmrpaAnWupni7DuoTfb4vXRT69WHhjwChS-GLa27qAvg6BvZl4kChO6O5STiDRny7f9ebFE41Dr8jFL7EOtL0tFIJy5azlOtdH-gtU-uiIpejggBULinh2qsP8aM6ywu6ozbIsRx4D77FFLRsyLwc5hjWYn8S-M-_9H6_b1Z92oSai0HIkUEWn2k72UIH9kTwxK88tYzOWhwC9yrcHaBs33sleQcCInhytk3VUd0UmWDwz659y-mjVq5SbM4dChkjTfI0TIagxulPU8LedKwJ2uLHBKOgZx9N2fZHpiBje4S0iYrFdoC4y_61tJ-0T209rIeLEM3L9FTQsVkxlGqNnPNwe2ocVaLVehn4DSJZ_qmWFe4VrkBc6oGwnD-5GaqnlBSCVII3ex94mHpSBXHpmZRZxI1gJx7xjECAlEu-8LQ45vdqWZ5osSp3a1GN31xf2dWguwPMIUuTLAyEViBjxsMX4QDfyHTgQ7tGMAZDp3ELH49FKjPCL2bO1unawOf_h42fjeImqq2xAZVP_fQM6CuWLQ9CMD1jWyGnNOWj2foukVnESDsfQ7h_NKztFludRyV-eN8E4sXBIxPyzgOi-_XTx7pqE0cbiuXSsGIl_i2VJx2hjrINPPuf0RlRc2IzHhu9-WfVTwbC5IKLPerfNQz6qjeKDvq8qjNJFbj51jSWnVkeQEqulrf7iF7CJ_d3xAGFfZzWk1zj0_VAxgWdxJCSSTOMPr4FdXAfICAJ3qBEklZFSkJXoEpc8KLhkOBBtbAD3HIGsKS_JE9-bOMfvrDHo6xMpeNPB3WrHI5FtPjhIGw.axRwkUzM8ZIHrj0u_xpBKg
import { getUserProfile } from "../helpers/sso-helper.js";
import { getCorrespondancesData } from "../helpers/lists-helper.js";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    document.getElementById("getProfileButton").onclick = run;
    document.getElementById("societesButton").onclick = societesClick;
    document.getElementById("correspondancesButton").onclick = correspondancesClick;
    document.getElementById("intervenantsButton").onclick = intervenantsClick;
    document.getElementById("searchInput").onchange = searchInputChange;
    document.getElementById("filterButton").onclick = filterButtonClick;
  }
});

var societesData;
var correspondancesData;
var intervenantsData;

function searchInputChange(value) {}

function filterButtonClick() {}

export async function run() {
  await getUserProfile(showHomeTaskPane);
}

function printData(data, type) {
  var html;
  var count;
  var btn2;
  var btn3;
  if (type == "societes") {
    html = "<table class='data-display'><th>id</th><th>Type</th><th>Raison Sociale</th>";
    count = 2;
    btn2 = "correspondants";
    btn3 = "intervenants";
  }
  if (type == "correspondances") {
    html = "<table class='data-display'><th>id</th><th>Nom</th><th>Prenom</th><th>Poste</th><th>email</th>";
    count = 4;
    btn2 = "societe";
    btn3 = "intervenants";
  }

  data.map((v) => {
    html += "<tr><td>" + v.id + "</td>";

    for (var i = 1; i <= count; i++) {
      html += "<td>" + v.fields["field_" + i] + "</td>";
    }
    html +=
      "<td><button>detail</button></td><td><button>" +
      btn2 +
      "</button></td><td><button>" +
      btn3 +
      "</button></td></tr>";
  });

  html += "</table>";
  return html;
}

async function showHomeTaskPane(data) {
  var pane = document.getElementById("home");
  document.getElementById("main").style.display = "none";
  document.getElementById("header").style.display = "none";
  pane.style.display = "block";

  console.log(data);
  societesData = data;

  document.getElementById("societes").innerHTML = printData(data, "societes");

  correspondancesData = await getCorrespondancesData();

  //Office.context.mailbox.item.body.setSelectedDataAsync(userInfo, { coercionType: Office.CoercionType.Html });
}

function societesClick() {
  document.getElementById("societes").style.display = "block";
  document.getElementById("correspondances").style.display = "none";
  document.getElementById("intervenants").style.display = "none";
}

function correspondancesClick() {
  document.getElementById("societes").style.display = "none";
  document.getElementById("correspondances").style.display = "block";
  document.getElementById("intervenants").style.display = "none";
  document.getElementById("correspondances").innerHTML = printData(correspondancesData, "correspondances");
}

function intervenantsClick() {
  document.getElementById("societes").style.display = "none";
  document.getElementById("correspondances").style.display = "none";
  document.getElementById("intervenants").style.display = "block";
}
