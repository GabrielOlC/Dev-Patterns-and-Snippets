/* -------------------------------------------------------------------------
   Source: FEBRACA Forms, page cadastro-afiliado (wf_afiliado.js)
   Name: cm_address
   Purpose: Looks up an address from a CEP through a chain of public
       providers, and always answers in the same shape, whichever provider
       replied.
   Context: Filled street, district, city and state on an affiliation form as
       soon as the eighth CEP digit was typed.
   Technique: Providers are tried in order. The chain moves on after an
       'error' and after a 'not_found', since the providers' coverage differs;
       the first 'not_found' is remembered, so the final verdict says "does
       not exist" rather than "failed". Adding a provider is one entry in
       dtProviders: a URL builder and a normaliser.
   Requirements:
       - AwesomeAPI CEP (https://cep.awesomeapi.com.br) and ViaCEP
         (https://viacep.com.br): public APIs, no key, called from the browser.

   Parameters:
       - bv_dtOptions: optional, and so is each key; (bv_vCEP, bv_vsCallback)
         also works:
           { arProviders: ['awesomeapi', 'viacep'], vTimeoutMs: 8000 }
         arProviders is the order to try (arDefaultChain when omitted);
         vTimeoutMs caps each provider, not the whole lookup.
       - bv_vsCallback: receives one address, in the same shape from every
         provider:
           {
             vStatus:   'ok' | 'not_found' | 'invalid' | 'error',
             vProvider: 'awesomeapi',        // who answered; '' when none did
             vCEP:      '01310100',          // digits only, as looked up
             vStreet:   'Avenida Paulista',
             vDistrict: 'Bela Vista',
             vCity:     'São Paulo',
             vState:    'SP',
             vLat:      '-23.5632188',       // '' when the provider has none
             vLng:      '-46.6542596'
           }
   Notes: vStatus 'invalid' means not 8 digits, and no request is made;
       'error' means a provider did not answer, or answered something
       unexpected. Every field is always present, as a trimmed string.
       The request goes from the browser straight to a third party, so the
       provider sees the visitor's IP. Under a Content-Security-Policy, both
       domains must be in connect-src, or the lookup fails silently.

   Version: 261030
   ------------------------------------------------------------------------- */

(function (global) {
  'use strict';

  var cDigitsCEP        = 8;
  var cDefaultTimeoutMs = 8000;

  /* AwesomeAPI first, for its coordinates. */
  var arDefaultChain = ['awesomeapi', 'viacep'];

  /* ---------- Helpers ------------------------------------------------------ */

  function fnOnlyDigits(bv_vValue) {
    return String(bv_vValue === null || bv_vValue === undefined ? '' : bv_vValue)
      .replace(/\D+/g, '');
  }

  /* Third-party text: only ever a trimmed string, never null in a field. */
  function fnText(bv_vValue) {
    return typeof bv_vValue === 'string' ? bv_vValue.trim() : '';
  }

  /* The normalised answer, always with every field present. */
  function fnResult(bv_vStatus, bv_vProvider, bv_vDigits, bv_dtFields) {
    var dtOut = {
      vStatus:   bv_vStatus,
      vProvider: bv_vProvider || '',
      vCEP:      bv_vDigits || '',
      vStreet:   '',
      vDistrict: '',
      vCity:     '',
      vState:    '',
      vLat:      '',
      vLng:      ''
    };

    if (bv_dtFields) {
      for (var vKey in bv_dtFields) {
        if (Object.prototype.hasOwnProperty.call(bv_dtFields, vKey)) {
          dtOut[vKey] = fnText(bv_dtFields[vKey]);
        }
      }
    }
    return dtOut;
  }

  /* ---------- Providers ----------------------------------------------------
     Each entry: how to build the URL, and how to read the answer. The shapes
     below were checked against the real APIs.
     ---------------------------------------------------------------------- */

  var dtProviders = {

    /* https://cep.awesomeapi.com.br/json/01310100
       ok    -> { cep, address_type, address_name, address, state, district,
                  lat, lng, city, city_ibge, ddd }
       error -> HTTP 404 { code: 'not_found', message: '...' }             */
    awesomeapi: {
      vLabel: 'AwesomeAPI',
      vHasCoords: true,
      fnURL: function (bv_vDigits) {
        return 'https://cep.awesomeapi.com.br/json/' + bv_vDigits;
      },
      fnNormalize: function (bv_vHTTPStatus, bv_dtBody, bv_vDigits) {
        if (bv_vHTTPStatus === 404) {
          return fnResult('not_found', 'awesomeapi', bv_vDigits);
        }
        if (bv_vHTTPStatus < 200 || bv_vHTTPStatus >= 300 || !bv_dtBody || !bv_dtBody.cep) {
          return fnResult('error', 'awesomeapi', bv_vDigits);
        }
        return fnResult('ok', 'awesomeapi', bv_vDigits, {
          vStreet:   bv_dtBody.address,
          vDistrict: bv_dtBody.district,
          vCity:     bv_dtBody.city,
          vState:    bv_dtBody.state,
          vLat:      bv_dtBody.lat,
          vLng:      bv_dtBody.lng
        });
      }
    },

    /* https://viacep.com.br/ws/01310100/json/
       ok    -> { cep: '01310-100', logradouro, complemento, bairro,
                  localidade, uf, estado, regiao, ibge, ... }
       error -> HTTP **200** with { erro: 'true' }, a STRING or a boolean
                depending on the version; both are accepted below.
       No coordinates.                                                     */
    viacep: {
      vLabel: 'ViaCEP',
      vHasCoords: false,
      fnURL: function (bv_vDigits) {
        return 'https://viacep.com.br/ws/' + bv_vDigits + '/json/';
      },
      fnNormalize: function (bv_vHTTPStatus, bv_dtBody, bv_vDigits) {
        if (bv_vHTTPStatus < 200 || bv_vHTTPStatus >= 300 || !bv_dtBody) {
          return fnResult('error', 'viacep', bv_vDigits);
        }
        if (bv_dtBody.erro === true || bv_dtBody.erro === 'true') {
          return fnResult('not_found', 'viacep', bv_vDigits);
        }
        if (!bv_dtBody.cep) {
          return fnResult('error', 'viacep', bv_vDigits);
        }
        return fnResult('ok', 'viacep', bv_vDigits, {
          vStreet:   bv_dtBody.logradouro,
          vDistrict: bv_dtBody.bairro,
          vCity:     bv_dtBody.localidade,
          vState:    bv_dtBody.uf
        });
      }
    }
  };

  /* ---------- One provider ------------------------------------------------- */

  function vsAskProvider(bv_vId, bv_vDigits, bv_vTimeoutMs, bv_vsCallback) {
    var dtProvider = dtProviders[bv_vId];
    if (!dtProvider) {
      return bv_vsCallback(fnResult('error', '', bv_vDigits));
    }

    var sController = global.AbortController ? new global.AbortController() : null;
    var vIsDone = false;
    var vTimer  = global.setTimeout(function () {
      if (sController) { sController.abort(); }
    }, bv_vTimeoutMs);

    function vsFinish(bv_dtOut) {
      if (vIsDone) { return; }
      vIsDone = true;
      global.clearTimeout(vTimer);
      bv_vsCallback(bv_dtOut);
    }

    global.fetch(dtProvider.fnURL(bv_vDigits), {
      method: 'GET',
      /* No headers of our own: any of them would bring an OPTIONS preflight. */
      mode: 'cors',
      signal: sController ? sController.signal : undefined
    }).then(function (bv_sResponse) {
      var vHTTPStatus = bv_sResponse.status;
      return bv_sResponse.json().then(function (bv_dtBody) {
        vsFinish(dtProvider.fnNormalize(vHTTPStatus, bv_dtBody, bv_vDigits));
      }, function () {
        /* A body that is not JSON: a 404 like that still means "not found";
           anything else is the provider failing. */
        vsFinish(fnResult(vHTTPStatus === 404 ? 'not_found' : 'error',
          bv_vId, bv_vDigits));
      });
    })['catch'](function () {
      vsFinish(fnResult('error', bv_vId, bv_vDigits));
    });
  }

  /* ---------- The lookup, along the chain of providers --------------------- */

  function vsLookupByCEP(bv_vCEP, bv_dtOptions, bv_vsCallback) {
    /* Tolerant signature: (cep, callback) works too. */
    if (typeof bv_dtOptions === 'function') {
      bv_vsCallback = bv_dtOptions;
      bv_dtOptions = null;
    }
    var dtOptions = bv_dtOptions || {};

    var vDigits = fnOnlyDigits(bv_vCEP);

    /* An incomplete CEP is nobody's failure, and spends no request. */
    if (vDigits.length !== cDigitsCEP) {
      return bv_vsCallback(fnResult('invalid', '', vDigits));
    }

    var arChain = dtOptions.arProviders && dtOptions.arProviders.length
      ? dtOptions.arProviders
      : arDefaultChain;
    var vTimeoutMs = dtOptions.vTimeoutMs || cDefaultTimeoutMs;

    var vIndex = 0;
    var dtNotFound = null;

    function vsNext() {
      if (vIndex >= arChain.length) {
        return bv_vsCallback(dtNotFound || fnResult('error', '', vDigits));
      }

      var vId = arChain[vIndex++];
      vsAskProvider(vId, vDigits, vTimeoutMs, function (bv_dtOut) {
        if (bv_dtOut.vStatus === 'ok') { return bv_vsCallback(bv_dtOut); }
        if (bv_dtOut.vStatus === 'not_found' && !dtNotFound) { dtNotFound = bv_dtOut; }
        vsNext();
      });
    }

    vsNext();
  }

  /* ---------- Outside the lookup ------------------------------------------- */

  /* "Rua X, 100, sala 4 — Bairro". Empty parts drop out. */
  function fnFormatOneLine(bv_dtAddress, bv_vNumber, bv_vComplement) {
    var vStreet   = fnText(bv_dtAddress && bv_dtAddress.vStreet);
    var vNumber   = fnText(bv_vNumber);
    var vExtra    = fnText(bv_vComplement);
    var vDistrict = fnText(bv_dtAddress && bv_dtAddress.vDistrict);

    var vLine = vStreet;
    if (vNumber)   { vLine += (vLine ? ', ' : '') + vNumber; }
    if (vExtra)    { vLine += (vLine ? ', ' : '') + vExtra; }
    if (vDistrict) { vLine += (vLine ? ' — ' : '') + vDistrict; }
    return vLine;
  }

  global.cm_address = {
    cDigitsCEP: cDigitsCEP,
    arDefaultChain: arDefaultChain,
    fnFormatOneLine: fnFormatOneLine,
    vsLookupByCEP: vsLookupByCEP
  };

}(window));


/* -------------
   Usage Example
   -------------

cm_address.vsLookupByCEP('01310-100', { vTimeoutMs: 5000 }, function (dtAddress) {
  if (dtAddress.vStatus !== 'ok') { return; }
  document.getElementById('fmCity').value = dtAddress.vCity;   // 'São Paulo'
});

*/
