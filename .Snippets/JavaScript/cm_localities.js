/* -------------------------------------------------------------------------
   Source: FEBRACA Forms, page cadastro-afiliado (wf_afiliado.js)
   Name: cm_localities
   Purpose: Lists Brazil's 27 federative units and each one's municipalities,
       and turns what a person typed into the official IBGE spelling, or ''
       when it is not a municipality of that state.
   Context: Fed the state select and the city suggestions on an affiliation
       form, and made the city field accept only real municipalities.
   Technique: The 27 UFs are inline, because they are 27 and never change.
       The 5,571 municipalities are one generated file per UF, and only the
       chosen state's file is loaded, through a <script> tag rather than
       fetch, so it also works when the page is opened from disk. Each file
       registers itself as it parses. Names are matched ignoring accents, case
       and repeated spaces.
   Requirements:
       - ../data/municipalities/<UF>.js: one file per UF, each calling
         cm_localities.vsRegisterMunicipalities('<UF>', [...]). The path is
         resolved from this file's own URL, so js/ and data/ sit side by side,
         as in a project's shared/.
       - cm_municipalities.py writes those files from the IBGE API
         (python cm_municipalities.py <project>/shared/data/municipalities);
         they are never edited by hand.

   Parameters:
       - bv_vUF: the UF's two letters, in any case.
       - bv_vsCallback (vsLoadUF): function (sError, arNames); sError is null
         on success. Calls made while the file is still loading wait for it.
   Notes: A file that loads without registering itself is reported as an
       error. fnFormatCityUF builds a single 'Cidade/UF' value, for a schema
       that keeps both in one field.

   Attribution:
       - Original Concept: both lists come from IBGE, API de Localidades v1
         (https://servicodados.ibge.gov.br/api/docs/localidades), never typed
         by hand.

   Version: 261030
   ------------------------------------------------------------------------- */

(function (global) {
  'use strict';

  /* Resolved from this file's own URL so the data path holds however the site
     is mounted, and works when the page is opened straight off disk. */
  var cBasePath = (function () {
    var sScript = document.currentScript;
    if (!sScript || !sScript.src) { return ''; }
    return sScript.src.replace(/[^/]*$/, '');
  }());

  var cDataPath = cBasePath + '../data/municipalities/';

  var arUFs = [
    { vUF: 'AC', vName: 'Acre' },
    { vUF: 'AL', vName: 'Alagoas' },
    { vUF: 'AM', vName: 'Amazonas' },
    { vUF: 'AP', vName: 'Amapá' },
    { vUF: 'BA', vName: 'Bahia' },
    { vUF: 'CE', vName: 'Ceará' },
    { vUF: 'DF', vName: 'Distrito Federal' },
    { vUF: 'ES', vName: 'Espírito Santo' },
    { vUF: 'GO', vName: 'Goiás' },
    { vUF: 'MA', vName: 'Maranhão' },
    { vUF: 'MG', vName: 'Minas Gerais' },
    { vUF: 'MS', vName: 'Mato Grosso do Sul' },
    { vUF: 'MT', vName: 'Mato Grosso' },
    { vUF: 'PA', vName: 'Pará' },
    { vUF: 'PB', vName: 'Paraíba' },
    { vUF: 'PE', vName: 'Pernambuco' },
    { vUF: 'PI', vName: 'Piauí' },
    { vUF: 'PR', vName: 'Paraná' },
    { vUF: 'RJ', vName: 'Rio de Janeiro' },
    { vUF: 'RN', vName: 'Rio Grande do Norte' },
    { vUF: 'RO', vName: 'Rondônia' },
    { vUF: 'RR', vName: 'Roraima' },
    { vUF: 'RS', vName: 'Rio Grande do Sul' },
    { vUF: 'SC', vName: 'Santa Catarina' },
    { vUF: 'SE', vName: 'Sergipe' },
    { vUF: 'SP', vName: 'São Paulo' },
    { vUF: 'TO', vName: 'Tocantins' }
  ];

  var dtMunicipalities = {};   /* UF -> array of canonical names      */
  var dtIndex          = {};   /* UF -> { normalisedName: canonical } */
  var dtLoadState      = {};   /* UF -> loading | loaded | error      */
  var dtPending        = {};   /* UF -> array of callbacks            */

  function fnIsValidUF(bv_vUF) {
    var vUF = String(bv_vUF || '').toUpperCase();
    return arUFs.some(function (bv_dtUF) { return bv_dtUF.vUF === vUF; });
  }

  function fnUFName(bv_vUF) {
    var vUF = String(bv_vUF || '').toUpperCase();
    for (var z = 0; z < arUFs.length; z++) {
      if (arUFs[z].vUF === vUF) { return arUFs[z].vName; }
    }
    return '';
  }

  /* Accent- and case-insensitive key: "sao paulo" means São Paulo. */
  function fnNormalizeKey(bv_vName) {
    return String(bv_vName || '')
      .trim()
      .toLowerCase()
      .normalize('NFD')
      .replace(new RegExp('[\\u0300-\\u036f]', 'g'), '')
      .replace(new RegExp('[\\u2018\\u2019\\u0060\\u00b4]', 'g'), "'")
      .replace(/\s+/g, ' ');
  }

  function vsBuildIndex(bv_vUF) {
    var dtLookup = {};
    var arNames = dtMunicipalities[bv_vUF] || [];
    for (var z = 0; z < arNames.length; z++) {
      dtLookup[fnNormalizeKey(arNames[z])] = arNames[z];
    }
    dtIndex[bv_vUF] = dtLookup;
  }

  function vsFlush(bv_vUF, bv_sError) {
    var arCallbacks = dtPending[bv_vUF] || [];
    dtPending[bv_vUF] = [];
    for (var z = 0; z < arCallbacks.length; z++) {
      arCallbacks[z](bv_sError || null, bv_sError ? null : dtMunicipalities[bv_vUF]);
    }
  }

  /* Called by each generated data file as it parses. */
  function vsRegisterMunicipalities(bv_vUF, bv_arNames) {
    var vUF = String(bv_vUF).toUpperCase();
    dtMunicipalities[vUF] = bv_arNames;
    vsBuildIndex(vUF);
    dtLoadState[vUF] = 'loaded';
    vsFlush(vUF, null);
  }

  /* A <script> tag rather than fetch(): the data has to load when the page is
     opened from the file system as well as over HTTP, and fetch is blocked by
     CORS on file://. Same-origin static asset either way — no API is called. */
  function vsLoadUF(bv_vUF, bv_vsCallback) {
    var vUF = String(bv_vUF || '').toUpperCase();

    if (!fnIsValidUF(vUF)) {
      bv_vsCallback(new Error('UF desconhecida: ' + vUF), null);
      return;
    }

    if (dtLoadState[vUF] === 'loaded') {
      bv_vsCallback(null, dtMunicipalities[vUF]);
      return;
    }

    if (!dtPending[vUF]) { dtPending[vUF] = []; }
    dtPending[vUF].push(bv_vsCallback);

    if (dtLoadState[vUF] === 'loading') { return; }
    dtLoadState[vUF] = 'loading';

    var sScript = document.createElement('script');
    sScript.src = cDataPath + vUF + '.js';
    sScript.async = true;

    sScript.onload = function () {
      /* The file registers itself while parsing. Reaching onload without that
         having happened means the file loaded but was not what we expected. */
      if (dtLoadState[vUF] === 'loaded') { return; }
      dtLoadState[vUF] = 'error';
      vsFlush(vUF, new Error('Lista de municípios de ' + vUF + ' veio vazia.'));
    };

    sScript.onerror = function () {
      dtLoadState[vUF] = 'error';
      vsFlush(vUF, new Error('Não foi possível carregar os municípios de ' + vUF + '.'));
    };

    document.head.appendChild(sScript);
  }

  function fnIsLoaded(bv_vUF) {
    return dtLoadState[String(bv_vUF || '').toUpperCase()] === 'loaded';
  }

  function fnMunicipalities(bv_vUF) {
    return dtMunicipalities[String(bv_vUF || '').toUpperCase()] || [];
  }

  /* The canonical IBGE spelling for what was typed, or '' if it is not a
     municipality of that state. This is what makes the field constrained. */
  function fnCanonicalMunicipality(bv_vUF, bv_vName) {
    var vUF = String(bv_vUF || '').toUpperCase();
    var dtLookup = dtIndex[vUF];
    if (!dtLookup) { return ''; }
    return dtLookup[fnNormalizeKey(bv_vName)] || '';
  }

  function fnIsValidMunicipality(bv_vUF, bv_vName) {
    return fnCanonicalMunicipality(bv_vUF, bv_vName) !== '';
  }

  /* One combined value, 'Cidade/UF', for a schema that keeps both in a field. */
  function fnFormatCityUF(bv_vCity, bv_vUF) {
    var vCity = String(bv_vCity || '').trim();
    var vUF = String(bv_vUF || '').toUpperCase();
    if (!vCity || !vUF) { return ''; }
    return vCity + '/' + vUF;
  }

  global.cm_localities = {
    arUFs: arUFs,
    fnIsValidUF: fnIsValidUF,
    fnUFName: fnUFName,
    fnNormalizeKey: fnNormalizeKey,
    vsRegisterMunicipalities: vsRegisterMunicipalities,
    vsLoadUF: vsLoadUF,
    fnIsLoaded: fnIsLoaded,
    fnMunicipalities: fnMunicipalities,
    fnCanonicalMunicipality: fnCanonicalMunicipality,
    fnIsValidMunicipality: fnIsValidMunicipality,
    fnFormatCityUF: fnFormatCityUF
  };

}(window));


/* -------------
   Usage Example
   -------------

cm_localities.vsLoadUF('SP', function (sError) {
  if (sError) { return; }
  cm_localities.fnCanonicalMunicipality('SP', 'sao paulo');   // 'São Paulo'
});

*/
