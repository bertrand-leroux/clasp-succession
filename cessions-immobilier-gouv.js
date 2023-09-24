function fetchCessionsImmobilierGouv() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('cessions.immobilier-etat.gouv.fr');
  sheet.clear({ contentsOnly: true });
  const data = sheet.getDataRange().getValues();

  const baseUrl = "https://cessions.immobilier-etat.gouv.fr/index.php";
  const url = baseUrl + "/recherche?ajax_form=1&_wrapper_format=drupal_ajax";
  var formData = {
    "form_id": "formulaire_recherche",
    "_triggering_element_name": "rechercher",
    "tri_biens": "tri_proximite",
    "type_env_transports_recherche": "on",
    "type_env_education_recherche": "on",
    "type_env_commerces_recherche": "on",
    "type_env_sante_recherche": "on",
    "type_env_services_recherche": "on",
    //"type_maison": "on",
    // "type_appartement": "on",
    // "type_hotel": "on",
    // "type_hotel_particulier": "on",
    // "type_immeuble": "on",
    // "type_chateau": "on",
    // "type_parking": "on",
    // "type_cave": "on",
    // "type_logement_ensembles_immo": "on",
    // "type_terrain_nu": "on",
    // "type_terrain_occupe": "on",
    // "type_terrain_agricole": "on",
    // "type_terrain_boise": "on",
    // "type_vigne_verger": "on"
  };
  var options = {
    'method': 'post',
    'payload': formData
  };
  let response = UrlFetchApp.fetch(url, options);
  let content = response.getContentText();

  let json = JSON.parse(content)
  let points = json[2].liste_points;

  sheet.appendRow([
    `num_dossier`,
    `type_bien`,
    `titre`,
    `localisation`,
    `addresse`,
  ]);
  const baseMapsUrl = "https://www.google.com/maps/search/";
  let line = 2;
  Object.entries(points).forEach((entry) => {
    const [key, value] = entry;
    let itemUrl = baseUrl + value.url;
    let mapsUrl = `${baseMapsUrl + value.lat},+${value.lng}`;
    let itemContent = UrlFetchApp.fetch(itemUrl).getContentText();
    const $ = Cheerio.load(itemContent);
    let address = $('#bien-tab-map').attr('data-popup-text');
    const mapsSearchUrl = 'https://www.google.com/maps/search/?api=1&query=';

    sheet.appendRow([
      value.num_dossier,
      value.type_bien,
      `=HYPERLINK("${itemUrl}";"${value.titre}")`,
      `=HYPERLINK("${mapsUrl}";"${value.localisation}")`,
      `=HYPERLINK("${mapsSearchUrl + address}";"${address}")`,
    ]);
    line++;
  });

  SpreadsheetApp.getUi().alert(`Site cessions.immobilier-etat.gouv.fr\nNombre d'annonces récupérées : ${line - 2}`)


}
