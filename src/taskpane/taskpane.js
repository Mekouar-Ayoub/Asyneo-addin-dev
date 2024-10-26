/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global document, Office */
//loki auth token from session storage,
// eyJhbGciOiJSU0EtT0FFUCIsImVuYyI6IkExMjhDQkMtSFMyNTYiLCJ4NXQiOiJtRXpHeF80OE9MOWcxdXh1QkFnRXVOdGJqcVEiLCJ6aXAiOiJERUYifQ.J0GAJX7SnDfdsoXMboB4Vm8G4jTUs3aBPWCycFe4AlnK-LXxRdKmKDLswEDAetH_FK-M64kQzDmfbdqGR-BK2gFf9jsrA2T2QMJfOuYVH2mEKotr7MRnkYTRyPXpQ4xfyvjT3J2W9fNpF0OSENHcU872_fNXxqliOXpqQuM7tbDPmrdkdkDNWNFiDxFJaPYRJBY-ITXnbCUv7Q6V9pVS5XGUKvKKu5iW7Zy3IGRDpXaQz8tjAWQPhkyNmXRgibrzAqHa_C4PdlKuQg-RuJM69I-u1rcYXxvUDWTpsNme5fRaPGbg73eli_h6uux4bFsvBYVvnizA3Sjapxdc8x1oTg.IzNlswzsjOZ9M69rh3IpFA.xpgqFsFOLIpMh2brsqa4MIJUr0uR52QYHROeQSufglYeyWqE7aVUpqDDu6dV8J4RE1Elvc_5wXkMYCcoLJF_vIaZ8MW8EYMipdR8Cj4zSpSvLWa29U0NcT5_pEnw3qitzBGQm-xm0eUwmZJMoykOzpmaRHCrXtv-bCbuRqI9ra9nzMtnsQIfyEFTSP8QfSX4xvEfezQZoZHeuQWVpykygp_HMaf7xLnvQwmoTfn9nyfDO6S_E9qEu_EtsBjUo2JlfyQfL2ErlZ0raEu8NbVd5wCIldFMl2zkGZ4P6JUW3fLmKqnrn_SgoW7X47wPJ5V-SrMNlaAi3ljQvvnfeSSqU9374V87GT2xV7mRYud-TaTMneqq0DdF3ScGlHnbfsBWPxqg3XAZSyInNDGUkbbhuRLZc3akky4oGnM-LNDA4YG476Mk6E7fod8YEMIfy8D8GeXg6NO3agY2n46CBG55tzdoilqFAnwK40ml4PvuX5pvBfTA0_56CiAliVCbm5TTGuXr2X3PDfXNI1oHEnC3nfq5la8CUc9E64veNiM66Rebefy2m3KSJQ2yHI8f7VHHxRtV7y_skEXxLmsP87JJLGQ_vnb366hJjfa7iMaOWo0l7hLfv4u4fq_JzkZsL96AKPbwnDtY6W8eBjmdHLO3rs72YzPDmZNMm7JzwRtE3S2eo2QLEqnQKHCCRPkORSr0r3i7fOw_4A5CUU35tMyBtHgVOzkRMQ-Vcjf7HIv5-7sJMF2-d1LiadFMJVKKtmwd3PtqwtR1VPMC6DUpPKhFoHyHdf4t7EQIH5litSfCsZG7UfqE_R4qw5X_kIFM8_kO13CcxMY-Oxlw5kCQT7CWrpsuWIO6I9kRYImrYohmrpaAnWupni7DuoTfb4vXRT69WHhjwChS-GLa27qAvg6BvZl4kChO6O5STiDRny7f9ebFE41Dr8jFL7EOtL0tFIJy5azlOliH-gtU-uiIpejggBULinh2qsP8aM6ywu6ozbIsRx4D77FFLRsyLwc5hjWYn8S-M-_9H6_b1Z92oSai0HIkUEWn2k72UIH9kTwxK88tYzOWhwC9yrcHaBs33sleQcCInhytk3VUd0UmWDwz659y-mjVq5SbM4dChkjTfI0TIagxulPU8LedKwJ2uLHBKOgZx9N2fZHpiBje4S0iYrFdoC4y_61tJ-0T209rIeLEM3L9FTQsVkxlGqNnPNwe2ocVaLVehn4DSJZ_qmWFe4VrkBc6oGwnD-5GaqnlBSCVII3ex94mHpSBXHpmZRZxI1gJx7xjECAlEu-8LQ45vdqWZ5osSp3a1GN31xf2dWguwPMIUuTLAyEViBjxsMX4QDfyHTgQ7tGMAZDp3ELH49FKjPCL2bO1unawOf_h42fjeImqq2xAZVP_fQM6CuWLQ9CMD1jWyGnNOWj2foukVnESDsfQ7h_NKztFludRyV-eN8E4sXBIxPyzgOi-_XTx7pqE0cbiuXSsGIl_i2VJx2hjrINPPuf0RlRc2IzHhu9-WfVTwbC5IKLPerfNQz6qjeKDvq8qjNJFbj51jSWnVkeQEqulrf7iF7CJ_d3xAGFfZzWk1zj0_VAxgWdxJCSSTOMPr4FdXAfICAJ3qBEklZFSkJXoEpc8KLhkOBBtbAD3HIGsKS_JE9-bOMfvrDHo6xMpeNPB3WrHI5FtPjhIGw.axRwkUzM8ZIHrj0u_xpBKg
import { getUserProfile } from "../helpers/sso-helper.js";
import { getCorrespondancesData, getCorrespondances_telData, getSociete_telData, getIntervenantsData, getIntervenants_telData} from "../helpers/lists-helper.js";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);
    updateTaskPaneUI(Office.context.mailbox.item);

    // Attach functions without calling immediately
    document.getElementById("getProfileButton").onclick = run;
    document.getElementById("societesButton").onclick = societesClick;
    document.getElementById("correspondancesButton").onclick = correspondancesClick;
    document.getElementById("intervenantsButton").onclick = intervenantsClick;
    document.getElementById("filter-value").onchange = filterButtonClick;
    document.getElementById("filterButton").onclick = filterButtonClick;
    document.getElementById("routes-title").onclick = () => {
      document.getElementById("home").style.display = "flex";
    };
    const elements = document.querySelectorAll(".animation-chek");
    elements.forEach((element) => {
      element.addEventListener("click", animationchek);
    });
  }
});

var societesData;
var correspondancesData;
var correspondances_telData;
var societes_telData;
var intervenantsData;
var intervenants_telData;

function useState(initialValue) {
  let state = initialValue;

  function setFilter(newValue) {
    state = newValue;
    console.log("Updated state:", state);  // Optional: Log the new state
  }

  function filter() {
    return state;
  }

  return [filter, setFilter];
}

const [filter, setFilter] = useState('');
const [page, setpage] = useState('societe');
const [client, setClient] = useState(false);
const [sortie, setSortie] = useState(false);

function animationchek(event) {
  const container = event.currentTarget;  // Use event.currentTarget to target the clicked ".animation-chek" element
  const circle = container.querySelector(".circle");  // Find the ".circle" inside the clicked ".animation-chek"

  // Toggle the right position and background color
  if (circle.style.right === "50%") {
    circle.style.right = "0%";
    circle.style.backgroundColor = "#0072C6";
  } else{
    circle.style.right = "50%";
    circle.style.backgroundColor = "gray";
  }
  console.log("successfully clicked");
  updateTaskPaneUI(Office.context.mailbox.item);
}

function itemChanged(eventArgs) {
  
  updateTaskPaneUI(Office.context.mailbox.item);
}

function updateTaskPaneUI(item) {
  // Assuming that item is always a read item (instead of a compose item).
  if (item != null) console.log(item);

  // // TODO adapt based on item properties
  // if(item.subject == ) {
  //   //msgRead
  // }

  // if(item.subject == ) {
  //   //mdgCOmpose
  // }
}


function searchInputChange(value) {}

function filterButtonClick() {
  if(page()=="societe"){
    const filterValue = document.getElementById("filter-value").value; // Get input value
    setFilter(filterValue); // Update state with the filter value
    showHomeTaskPane(societesData); 
  }
  if(page()=="intervenant"){
    const filterValue = document.getElementById("filter-value").value; // Get input value
    setFilter(filterValue); // Update state with the filter value
    intervenantsClick(); 
  }
  if(page()=="correspondance"){
    const filterValue = document.getElementById("filter-value").value; // Get input value
    setFilter(filterValue); // Update state with the filter value
    correspondancesClick(); 
  }
}
export async function run() {
  await getUserProfile(showHomeTaskPane);
  await new Promise((resolve) => {
    setTimeout(() => {
      document.getElementById("body").style.overflow = "auto";
      resolve();
    }, 4000);
  });
}

function printDetails(element) {
  let html = "<ul class='details-container'>";

  // Loop over each property in `element.fields`
  for (const [key, value] of Object.entries(element.fields)) {
    html += `<li><strong>${key}:</strong> ${value}</li>`;
  }

  html += "</ul>";

  // Display the generated HTML in a specific container, e.g., with id "details"
  document.getElementById("societe-details-container").innerHTML = html;
  document.getElementById("societes").style.display ="none";
}

function printdata(data, type, tel, telType ,filter_content) {
  var html;
  var count;

  if (type === "societes" && telType === "telephones_societes") {
    html = "<ul class='data-display'>";
      count = 2;
      data.map((v) => {
        if(filter_content){
            if(!filter_content || (v.fields["field_2"] && v.fields["field_2"].toLowerCase().includes(filter_content.toLowerCase()))){
                html +="<ul class='data-cart'>" + "<li class='societe-title'>"+"<span class='societe-name'>" + v.fields["field_2"] + "</span>"+"<span>ID: " + v.fields["Title"] + "</span>" +"</li>" +  
                    "<li class='societe-ville'>"+"<img src='../../assets/home-icon.png'>"  + v.fields["field_15"] + "</li>";
                const matchedTel = tel.find(e => e.fields["field_3"] == v.fields["Title"]);
                if (matchedTel) {
                  html += "<li class='societe-telephone'>"+"<img src='../../assets/telephone-icon.png'>" +" +"+ matchedTel.fields["field_2"] + "</li>";
                } else {
                  html += "<li class='societe-telephone'>No phone number</li>";
                }

                html += `<li class='button-details'><button id='detailsButton' onclick='printDetails(${JSON.stringify(v)})'>Voir details <img src='../../assets/icon-right-arrow.png'></button></li>`;;
                html += "</ul>";
              }
        }else{

              html +="<ul class='data-cart'>" + "<li class='societe-title'>"+"<span class='societe-name'>" + v.fields["field_2"] + "</span>"+"<span>ID: " + v.fields["Title"] + "</span>" +"</li>" + 
                      
                      "<li class='societe-ville'>"+"<img src='../../assets/home-icon.png'>"  + v.fields["field_15"] + "</li>";

              const matchedTel = tel.find(e => e.fields["field_3"] == v.fields["Title"]);
                if (matchedTel) {
                  html += "<li class='societe-telephone'>"+"<img src='../../assets/telephone-icon.png'>" +" +"+ matchedTel.fields["field_2"] + "</li>";
                } else {
                  html += "<li class='societe-telephone'>No phone number</li>";
                }

              html += `<li class='button-details'><button id='' onclick='printDetails(${JSON.stringify(v)})'>Voir details <img src='../../assets/icon-right-arrow.png'></button></li>`;
              html += "</ul>";
        }
      });
    
    html += "</ul>";

  }else if (type === "correspondances" && telType === "telephones_correspondances") {
    html = "<ul class='data-display'>";
    count = 4;

    data.map((v) => {
      if(filter_content){
        if(!filter_content || (v.fields["field_2"] && v.fields["field_2"].toLowerCase().includes(filter_content.toLowerCase()))){
          html += "<ul class='data-cart'>"+"<li>" + v.id + "</li>";
          for (var i = 1; i <= count; i++) {
            html += "<li>" + v.fields["field_" + i] + "</li>";
          }
    
          const matchedTel = tel.find(e => e.fields["field_3"] == v.fields["Title"]);
          if (matchedTel) {
            html += "<li>" + matchedTel.fields["field_2"] + "</li>";
          } else {
            html += "<li>No phone number</li>";
          }
    
          html += "<li class='button-details'><button>Voir details <img src='../../assets/icon-right-arrow.png'> </button></li>";
          html += "</ul>";
          }
    }else{
      html += "<ul class='data-cart'>"+"<li>" + v.id + "</li>";
      for (var i = 1; i <= count; i++) {
        html += "<li>" + v.fields["field_" + i] + "</li>";
      }

      const matchedTel = tel.find(e => e.fields["field_3"] == v.fields["Title"]);
      if (matchedTel) {
        html += "<li>" + matchedTel.fields["field_2"] + "</li>";
      } else {
        html += "<li>No phone number</li>";
      }

      html += "<li class='button-details'><button>Voir details <img src='../../assets/icon-right-arrow.png'> </button></li>";
      html += "</ul>";
      }
    });
    
    html += "</ul>";

  } else if (type === "intervenants" && telType === "telephones_intervenants") {
    html = "<ul class='data-display'>";
    count = 5;
    data.map((v) => {
      if(filter_content){
        if(!filter_content || (v.fields["field_2"] && v.fields["field_2"].toLowerCase().includes(filter_content.toLowerCase()))){
          html +="<ul class='data-cart'>" + "<li>" + v.id + "</li>";
          for (var i = 1; i <= count; i++) {
            html += "<li>" + v.fields["field_" + i] + "</li>";
          }
          const matchedTel = tel.find(e => e.fields["field_3"] == v.fields["Title"]);
          if (matchedTel) {
            html += "<li>" + matchedTel.fields["field_2"] + "</li>";
          } else {
            html += "<li>No phone number</li>";
          }

          html += "<li class='button-details'><button>Voir details <img src='../../assets/icon-right-arrow.png'></button></li>";
          html += "</ul>";
          }
    }else{
          html +="<ul class='data-cart'>" + "<li>" + v.id + "</li>";
          for (var i = 1; i <= count; i++) {
            html += "<li>" + v.fields["field_" + i] + "</li>";
          }
          const matchedTel = tel.find(e => e.fields["field_3"] == v.fields["Title"]);
          if (matchedTel) {
            html += "<li>" + matchedTel.fields["field_2"] + "</li>";
          } else {
            html += "<li>No phone number</li>";
          }

          html += "<li class='button-details'><button>Voir details <img src='../../assets/icon-right-arrow.png'></button></li>";
          html += "</ul>";
        }
        });
        
        html += "</ul>";
  }

  html += "</ul>";
  return html;
}


async function showHomeTaskPane(data) {
  
  document.getElementById("main").style.display = "none";
  document.getElementById("header").style.display = "none";

  console.log(data);
  societesData = data;

  [correspondancesData , societes_telData , correspondances_telData , intervenantsData , intervenants_telData ] = await Promise.all([
    getCorrespondancesData(),
    getSociete_telData(),
    getCorrespondances_telData(),
    getIntervenantsData(),
    getIntervenants_telData()
  ]);

  document.getElementById("societes").innerHTML = printdata(data, "societes" , societes_telData , "telephones_societes",filter());
  document.getElementById("route-state").innerHTML = "societes";
  document.getElementById("home").style.display = "none";
}

function societesClick() {
  setpage('societe');
  document.getElementById("societes").style.display = "block";
  document.getElementById("correspondances").style.display = "none";
  document.getElementById("intervenants").style.display = "none";
  document.getElementById("societe-details-container").style.display = "none";
  document.getElementById("route-state").innerHTML = "societes";
  document.getElementById("home").style.display = "none";
}

function correspondancesClick() {
  setpage('correspondance');
  document.getElementById("societes").style.display = "none";
  document.getElementById("correspondances").style.display = "block";
  document.getElementById("intervenants").style.display = "none";
  document.getElementById("correspondances").innerHTML = printdata(correspondancesData, "correspondances" , correspondances_telData , "telephones_correspondances" , filter());
  document.getElementById("route-state").innerHTML = "correspondances";
  document.getElementById("home").style.display = "none";
}



function intervenantsClick() {
  setpage('intervenant');
  document.getElementById("societes").style.display = "none";
  document.getElementById("correspondances").style.display = "none";
  document.getElementById("intervenants").style.display = "block";
  document.getElementById("intervenants").innerHTML = printdata(intervenantsData, "intervenants" , intervenants_telData , "telephones_intervenants" ,filter());
  document.getElementById("route-state").innerHTML = "intervenants";
  document.getElementById("home").style.display = "none";
}

