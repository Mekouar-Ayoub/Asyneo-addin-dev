import fetch from "node-fetch";

//TODO way to get bureautique id each time

async function fetchSharePointListByDisplayName(accessToken, listDisplayName) {
  // Use the filter parameter to get the list by display name
  const url = `c'${listDisplayName}'`;

  try {
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: "Bearer " + accessToken,
      },
    });

    console.log("statusCode:", response.status);

    if (!response.ok) {
      const errorBody = await response.text(); // Read the error response body for debugging
      throw new Error(`HTTP error! status: ${response.status}, body: ${errorBody}`);
    }

    const data = await response.json();
    console.log("Data:", data);

    // Check if the list exists and return it
    if (data.value && data.value.length > 0) {
      const listId = data.value[0].id; // Get the ID of the found list
      console.log(`Found list: ${data.value[0].displayName}`);
      return listId; // Return the ID to fetch items later
    } else {
      console.log(`List with display name '${listDisplayName}' not found.`);
      return null;
    }
  } catch (error) {
    console.error("Fetch error:", error);
  }
}

// Function to fetch items from a specific list
async function fetchListItems(accessToken, listId) {
  const url = `https://graph.microsoft.com/v1.0/sites/2beeb2a2-19f8-46cc-998b-d1fa94ba6c0f/lists/${listId}/items?$expand=fields`;

  try {
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: "Bearer " + accessToken,
      },
    });

    console.log("statusCode:", response.status);

    if (!response.ok) {
      const errorBody = await response.text(); // Read the error response body
      throw new Error(`HTTP error! status: ${response.status}, body: ${errorBody}`);
    }

    const items = await response.json();
    console.log("Items in the list:", items);
    return items.value;
  } catch (error) {
    console.error("Fetch error:", error);
  }
}

async function fetchListFields(accessToken, listId) {
  const url = `https://graph.microsoft.com/v1.0/sites/2beeb2a2-19f8-46cc-998b-d1fa94ba6c0f/lists/${listId}/fields`;

  try {
    const response = await fetch(url, {
      method: "GET",
      headers: {
        Authorization: "Bearer " + accessToken,
      },
    });

    console.log("statusCode:", response.status);

    if (!response.ok) {
      const errorBody = await response.text(); // Read the error response body
      throw new Error(`HTTP error! status: ${response.status}, body: ${errorBody}`);
    }

    const fields = await response.json();
    console.log("Fields in the list:", fields);

    return fields.value; // Return the field definitions
  } catch (error) {
    console.error("Fetch error:", error);
  }
}

export async function getOrganisationData(accessToken) {
  console.log(accessToken);
  localStorage.setItem("token", accessToken);
  return fetchSharePointListByDisplayName(accessToken, "societes").then((listId) => {
    if (listId) {
      return fetchListItems(accessToken, listId);
    }
  });
}

export async function getCorrespondancesData() {
  var accessToken = localStorage.getItem("token");

  return fetchSharePointListByDisplayName(accessToken, "correspondances").then((listId) => {
    if (listId) {
      return fetchListItems(accessToken, listId);
    }
  });
}

export async function getSociete_telData() {
  var accessToken = localStorage.getItem("token");

  return fetchSharePointListByDisplayName(accessToken, "telephones_societes").then((listId) => {
    if (listId) {
      return fetchListItems(accessToken, listId);
    }
  });
}


export async function getIntervenantsData() {
  var accessToken = localStorage.getItem("token");

  fetchSharePointListByDisplayName(accessToken, "intervenants").then((listId) => {
    if (listId) {
      return fetchListItems(accessToken, listId);
    }
  });
}
//let accessToken = localStorage.getItem("Token");

//accounts?$select=name,address1_city&$top=10
