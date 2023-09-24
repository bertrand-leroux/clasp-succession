function fetchAnnoncesMesInfos() {
  let sheet = getSheetByName('annonces.mesinfos.fr');
  let last = sheet.getLastRow();
  const dateOptions = { year: 'numeric', month: 'numeric', day: 'numeric' };

  const data = sheet.getDataRange().getValues();
  let lastDateValue = sheet.getRange(last, this.getColNumberByName(`published_at`, data)).getValue();

  let date = new Date(lastDateValue);
  date.setDate(date.getDate())
  let startDate = date.toLocaleDateString('fr-FR', dateOptions);

  let endDate =  new Date(lastDateValue);
  endDate.setMonth(endDate.getMonth() + 3)
  let endDateString = endDate.toLocaleDateString('fr-FR', dateOptions);
  


  const DOMAIN = 'https://annonces.mesinfos.fr';;
  const CONSULT_PATH = DOMAIN + '/consulter-domaine/liste_consulter';
  const NO_RESULT_SENTENCE = 'Aucune annonce ne correspond à votre recherche.';
  const CURRENT_PAGE_PARAMETER = 'current_page';
  const DATA_CONTAINER_SELECTOR = 'table > tbody';
  const DATA_ITEM_SELECTOR = DATA_CONTAINER_SELECTOR + ' > tr';

  const SUPPORTS = {
    "tous": 0,
    "affiches-parisiennes.com": 700000,
    "journal-du-btp.com": 700007,
    "le-tout-lyon.fr": 700006,
    "lemoniteur77.com": 700001,
    "lepatriote.fr": 700072,
    "lerepublicainduzes.fr": 700141,
    "lessor38.fr": 700004,
    "lessor42.fr": 700005,
    "nouvellespublications.com": 700003,
    "semaine-ile-de-france.fr": 700080,
    "tpbm-presse.com": 700002,
    "Affiches Parisiennes": 200002,
    "L'Essor Affiches Loire": 606178,
    "L'Essor Isere": 606179,
    "La Semaine de l'Ile-de-France": 200131,
    "Le Journal du B\u00e2timent et des Travaux Publics": 200271,
    "Le Moniteur de Seine et Marne": 200286,
    "Le Patriote Beaujolais": 200295,
    "Le R\u00e9gional": 200335,
    "Le R\u00e9publicain d'Uz\u00e8s": 200339,
    "Le Var Information": 200362,
    "Les Nouvelles Publications": 200407,
    "Tout Lyon": 200358,
    "TPBM - Semaine Provence": 200528,
    "Vaucluse Hebdo": 200533
  };

  let page = 1;

  let parameters = {
    search: '',
    category: '28',
    subcategory: '37',
    date_debut: startDate,
    date_fin: endDateString,
    slug: 'annonces-legales',
    support: SUPPORTS['tous'],
    current_page: page,
  };

  let options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(parameters),
  };

  let continueFetch = true;
  let rows = [];

  while (continueFetch) {
    let response = UrlFetchApp.fetch(CONSULT_PATH, options);
    let content = response.getContentText();

    const $ = Cheerio.load(content);
    let container = $(DATA_CONTAINER_SELECTOR);
    let containerText = container.text().trim();
    continueFetch = !(containerText.includes(NO_RESULT_SENTENCE)) && containerText.length > 0;

    let itemsSelector = $(DATA_ITEM_SELECTOR);
    if (continueFetch) {
      itemsSelector.each(function () {
        rows.push({
          savedAt: new Date().toLocaleDateString(),
          publishedAt: getPublishedAt($(this)),
          ref: getRef($(this)),
          civility: getCivility($(this)),
          fullName: getFullname($(this)),
          status: 'nouveau',
          gender: getGender($(this)),
          deceasedAt: getDeceasedAt($(this)),
          department: getDeparment($(this)),
          url: DOMAIN + getUrl($(this)),
          category: getCategory($(this)),
          subCategory: getSubcategory($(this)),
        })
      });
    }

    parameters.current_page = ++page;
    options.payload = JSON.stringify(parameters);
  }

let i = 0;
  rows.forEach(function (row) {
    sheet.appendRow(Object.values(row));
    ++i;
  })

  SpreadsheetApp.getUi().alert(`Site annonces.mesinfos.fr\nNombre de nouvelles annonces récupérées : ${i}`)
}

function getCategories(el) {
  return el.find(`td[data-type="Catégorie / Sous-catégorie : "]`).text();
}

function getCategory(el) {
  return getCategories(el).split("/")[0].trim();
}

function getSubcategory(el) {
  return getCategories(el).split("/")[1].trim();
}

function getPublishedAt(el) {
  return el.find("td[data-type='Date de Publication : ']").text().trim();
}

function getUrl(el) {
  return el.find("td[data-type='Société : '] > a").attr('href');
}

function getDeparment(el) {
  return el.find("td[data-type='Département : ']").text().trim();
}

function getDeceasedAt(el) {
  return getExcerptText(el).match(/.*décédée? le (\d{2}\/\d{2}\/\d{4}).*/)[1];
}

function getGender(el) {
  return { 'M.': 'H', 'Mme': 'F', 'Mlle': 'F', }[getCivility(el)] || '';
}

function getFullname(el) {
  return getIdentity(el).replace(getCivility(el), '').trim()
}

function getCivility(el) {
  return getIdentity(el).match(/(M\.|Mme|Mlle)/)[1] ?? null;
}

function getIdentity(el) {
  return el.find("td[data-type='Société : '] > a").text()
}

function getExcerptText(el) {
  return el.find("td[data-type='Extrait : '] > div.summary-popup > div.popup > div.content").text();
}

function getRef(el) {
  return getExcerptText(el).match(/Réf. ([-\w]+)/)[1] ?? null;
}
