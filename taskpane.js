let map;
let directionsService;
let directionsRenderer;

// Initialize Office Add-in Context
Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initAppControls();
  }
});

function initAppControls() {
  // Setup interface event listeners
  document.getElementById("btnAdd").addEventListener("click", () => addStopInput(""));
  document.getElementById("calc").addEventListener("click", calculateRoute);
  document.getElementById("btnFetchText").addEventListener("click", extractAddressFromEmail);
  document.getElementById("btnOpen").addEventListener("click", openInGoogleMapsExternal);

  // Initialize Drag and Drop sorting
  const el = document.getElementById('stopsList');
  if (el && typeof Sortable !== 'undefined') {
    Sortable.create(el, { animation: 150 });
  }

  // Pre-populate two empty stops to mimic starting point options
  addStopInput("");
  addStopInput("");

  // Safely trigger maps engine initialization if scripts loaded before office lifecycle
  if (typeof google !== 'undefined' && google.maps) {
    setupMapsEngine();
  }
}

window.initMapFallback = function() {
  if (Office && Office.context) {
    setupMapsEngine();
  }
};

function setupMapsEngine() {
  directionsService = new google.maps.DirectionsService();
  directionsRenderer = new google.maps.DirectionsRenderer();
  
  const defaultLatLng = { lat: 43.6532, lng: -79.3832 }; // Default view: Toronto region
  map = new google.maps.Map(document.getElementById("map"), {
    zoom: 7,
    center: defaultLatLng,
    disableDefaultUI: true,
    zoomControl: true
  });
  directionsRenderer.setMap(map);
}

function addStopInput(value = "") {
  const list = document.getElementById("stopsList");
  const li = document.createElement("li");
  li.className = "stop-item";
  
  const input = document.createElement("input");
  input.type = "text";
  input.className = "stop-input";
  input.placeholder = "Enter city, zip, or facility address";
  input.value = value;
  
  const removeBtn = document.createElement("button");
  removeBtn.className = "btn-remove";
  removeBtn.innerText = "✕";
  removeBtn.onclick = () => li.remove();

  li.appendChild(input);
  li.appendChild(removeBtn);
  list.appendChild(li);

  if (typeof google !== 'undefined' && google.maps.places) {
    new google.maps.places.Autocomplete(input);
  }
}

// Intercept selected text or full frame context securely from open Outlook item
function extractAddressFromEmail() {
  Office.context.mailbox.item.getSelectedDataAsync(Office.CoercionType.Text, (asyncResult) => {
    if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value.trim().length > 2) {
      populateFirstEmptyOrNewStop(asyncResult.value.trim());
    } else {
      // Fallback: If no text is specifically selected, scan full item text body context
      Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, (bodyResult) => {
        if (bodyResult.status === Office.AsyncResultStatus.Succeeded) {
          // Rudimentary pattern lookup matching address blocks or postal structures
          const text = bodyResult.value;
          const match = text.match(/\b\d+\s+[A-Za-z0-9\s,.]+ (?:Avenue|Ave|Street|St|Road|Rd|Blvd|Drive|Dr|Way|Lane|Ln)\b/i);
          if (match) {
            populateFirstEmptyOrNewStop(match[0]);
          }
        }
      });
    }
  });
}

function populateFirstEmptyOrNewStop(addressText) {
  const inputs = document.querySelectorAll(".stop-input");
  for (let input of inputs) {
    if (!input.value) {
      input.value = addressText;
      return;
    }
  }
  addStopInput(addressText);
}

function calculateRoute() {
  if (!directionsService) return;

  const inputs = document.querySelectorAll(".stop-input");
  const locations = Array.from(inputs).map(i => i.value.trim()).filter(v => v !== "");

  if (locations.length < 2) {
    document.getElementById("routeText").innerText = "Please provide at least 2 locations.";
    return;
  }

  const origin = locations[0];
  const destination = locations[locations.length - 1];
  const waypoints = [];

  for (let i = 1; i < locations.length - 1; i++) {
    waypoints.push({ location: locations[i], stopover: true });
  }

  directionsService.route({
    origin: origin,
    destination: destination,
    waypoints: waypoints,
    travelMode: google.maps.TravelMode.DRIVING,
    unitSystem: google.maps.UnitSystem.IMPERIAL
  }, (response, status) => {
    if (status === "OK") {
      directionsRenderer.setDirections(response);
      
      let totalMetres = 0;
      const legs = response.routes[0].legs;
      legs.forEach(leg => { totalMetres += leg.distance.value; });

      const totalMiles = totalMetres * 0.000621371;
      const ratePerMile = parseFloat(document.getElementById("rate").value) || 0;
      const currency = document.getElementById("currency").value;
      const finalCost = totalMiles * ratePerMile;

      document.getElementById("routeText").innerText = `Route: From ${legs[0].start_address.split(',')[0]} to ${legs[legs.length-1].end_address.split(',')[0]}`;
      document.getElementById("distance").innerText = `Distance: ${totalMiles.toFixed(1)} mi — ${currency} $${finalCost.toFixed(2)}`;
    } else {
      document.getElementById("routeText").innerText = "Routing failed: " + status;
    }
  });
}

function openInGoogleMapsExternal() {
  const inputs = document.querySelectorAll(".stop-input");
  const locations = Array.from(inputs).map(i => i.value.trim()).filter(v => v !== "");
  if (locations.length === 0) return;
  
  const url = `https://www.google.com/maps/dir/${locations.map(encodeURIComponent).join('/')}`;
  window.open(url, '_blank');
}