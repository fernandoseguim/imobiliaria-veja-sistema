var l=document.getElementById("location-header");

function getLocation() {
  if (navigator.geolocation){
    navigator.geolocation.getCurrentPosition(function(position) {

    var url = "http://nominatim.openstreetmap.org/reverse?lat="
    +position.coords.latitude+"&lon="
    +position.coords.longitude+"&format=json&json_callback=showLocation";
  
    var script = document.createElement('script');
    script.src = url;
    document.body.appendChild(script);

    },showError);
  }
  else{
      l.innerHTML="Seu browser não suporta Geolocalização.";
  }
}

function showLocation(response){
  const API_KEY = "AIzaSyBtVuaDkqvshbbEtGZR4qmja6WZu2lCabI"
  var location = response.address.city; 
  location += ", " + response.address.state; 
  l.innerHTML = location; 

}

function showError(error)
  {
  switch(error.code)
    {
    case error.PERMISSION_DENIED:
      l.innerHTML="Usuário rejeitou a solicitação de Geolocalização."
      break;
    case error.POSITION_UNAVAILABLE:
      l.innerHTML="Localização indisponível."
      break;
    case error.TIMEOUT:
      l.innerHTML="A requisição expirou."
      break;
    case error.UNKNOWN_ERROR:
      l.innerHTML="Algum erro desconhecido aconteceu."
      break;
    }
  }
