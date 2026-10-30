/* -------------------------------------------------------------------------
   Source: FEBRACA Forms, page cadastro-afiliado (wf_afiliado.js)
   Name: cm_validators
   Purpose: Validates form inputs: CNPJ (including the alphanumeric one), CPF,
       RG, Brazilian phone, e-mail, date of birth, whole numbers, website,
       Instagram and uploaded files. Every function takes its input as an
       argument and returns a verdict, so any form can call it.
   Context: Checked each field of the FEBRACA forms as it was typed and again
       on submit; the controller worded the message.
   Technique: CNPJ and CPF check digits use módulo 11. A CNPJ character is
       worth its ASCII code minus 48, so numeric and alphanumeric CNPJ share
       one arithmetic (SERPRO's rule, under Attribution).

   Parameters:
       - bv_vISODate: 'YYYY-MM-DD', as a date input gives it.
       - bv_vMinAge / bv_vMaxAge, bv_vMin / bv_vMax: inclusive bounds.
       - bv_sFile, bv_arAcceptedTypes: a File from an <input type="file">, and
         the MIME types it may have, e.g. ['image/png', 'image/jpeg'].
   Notes: Reads no DOM and writes no message. The fnCheck* functions return
       { vIsValid, vReason, … } instead of a boolean, so the caller can say
       what is wrong; each lists its reasons above it. RG is only a presence
       check (3+ characters), since its format varies by issuing state.

   Attribution:
       - Original Concept: CNPJ check digits as specified by SERPRO, "Cálculo
         dos dígitos verificadores do CNPJ alfanumérico"
         (https://serpro.gov.br/menu/noticias/videos/calculodvcnpjalfanaumerico.pdf)

   Version: 261030
   ------------------------------------------------------------------------- */

(function (global) {
  'use strict';

  function fnOnlyDigits(bv_vValue) {
    return String(bv_vValue === null || bv_vValue === undefined ? '' : bv_vValue)
      .replace(/\D+/g, '');
  }

  function fnTrim(bv_vValue) {
    return String(bv_vValue === null || bv_vValue === undefined ? '' : bv_vValue).trim();
  }

  /* ---------- CNPJ ------------------------------------------------------- */

  /* Fourteen characters: twelve of base, which may hold A-Z since 31 July 2026,
     plus two numeric check digits. Numeric CNPJ go through the same arithmetic. */

  var cCNPJPattern = /^[0-9A-Z]{12}[0-9]{2}$/;

  /* SERPRO's rule: the numeric value of a character is its ASCII code minus 48.
     '0' -> 0 … '9' -> 9, 'A' -> 17, 'B' -> 18 … 'Z' -> 42. */
  function fnCNPJCharValue(bv_vChar) {
    return bv_vChar.charCodeAt(0) - 48;
  }

  /* Uppercased on the way in: 'a' and 'A' are different ASCII codes and only one
     of them is the right answer. */
  function fnOnlyCNPJChars(bv_vValue) {
    return String(bv_vValue === null || bv_vValue === undefined ? '' : bv_vValue)
      .toUpperCase().replace(/[^0-9A-Z]+/g, '');
  }

  /* Módulo 11, weights cycling 2..9 from the right — the Receita Federal rule
     for both check digits. bv_vPosition is how many characters to the left are
     read: 12 for the first digit, 13 for the second. */
  function fnCNPJCheckDigit(bv_vChars, bv_vPosition) {
    var vSum = 0;
    var vWeight = 2;
    for (var z = bv_vPosition - 1; z >= 0; z--) {
      vSum += fnCNPJCharValue(bv_vChars.charAt(z)) * vWeight;
      vWeight = vWeight === 9 ? 2 : vWeight + 1;
    }
    var vRest = vSum % 11;
    return vRest < 2 ? 0 : 11 - vRest;
  }

  function fnIsValidCNPJ(bv_vValue) {
    var vChars = fnOnlyCNPJChars(bv_vValue);
    if (vChars.length !== 14) { return false; }

    /* Letters only in the base; the two check digits are numeric by rule. */
    if (!cCNPJPattern.test(vChars)) { return false; }

    /* A repeated character satisfies the arithmetic but is not a real
       registration — '00000000000000' checks out and does not exist. */
    if (/^(.)\1{13}$/.test(vChars)) { return false; }

    return fnCNPJCheckDigit(vChars, 12) === Number(vChars.charAt(12)) &&
           fnCNPJCheckDigit(vChars, 13) === Number(vChars.charAt(13));
  }

  /* ---------- CPF -------------------------------------------------------- */

  function fnCPFCheckDigit(bv_vDigits, bv_vPosition) {
    var vSum = 0;
    var vWeight = bv_vPosition + 1;   /* 10 for the first digit, 11 for the second */
    for (var z = 0; z < bv_vPosition; z++) {
      vSum += Number(bv_vDigits.charAt(z)) * vWeight;
      vWeight--;
    }
    var vRest = (vSum * 10) % 11;
    return vRest === 10 ? 0 : vRest;
  }

  function fnIsValidCPF(bv_vValue) {
    var vDigits = fnOnlyDigits(bv_vValue);
    if (vDigits.length !== 11) { return false; }
    if (/^(\d)\1{10}$/.test(vDigits)) { return false; }

    return fnCPFCheckDigit(vDigits, 9)  === Number(vDigits.charAt(9)) &&
           fnCPFCheckDigit(vDigits, 10) === Number(vDigits.charAt(10));
  }

  /* ---------- RG --------------------------------------------------------- */

  /* Deliberately only a presence check. The format varies by issuing state and
     has no national standard, so any pattern here produces false rejections. */
  function fnIsValidRG(bv_vValue) {
    return fnTrim(bv_vValue).length >= 3;
  }

  /* ---------- Telephone -------------------------------------------------- */

  /* Eleven digits is a mobile, whose subscriber number always opens with 9.
     Ten digits is a geographic landline, which opens 2-5. Area codes run
     11-99; there is no valid DDD below 11. */
  function fnIsValidPhoneBR(bv_vValue) {
    var vDigits = fnOnlyDigits(bv_vValue);
    if (vDigits.length !== 10 && vDigits.length !== 11) { return false; }

    var vAreaCode = Number(vDigits.slice(0, 2));
    if (vAreaCode < 11 || vAreaCode > 99) { return false; }

    if (vDigits.length === 11) { return vDigits.charAt(2) === '9'; }
    return /[2-5]/.test(vDigits.charAt(2));
  }

  function fnIsMobilePhoneBR(bv_vValue) {
    var vDigits = fnOnlyDigits(bv_vValue);
    return vDigits.length === 11 && fnIsValidPhoneBR(vDigits);
  }

  /* ---------- E-mail ----------------------------------------------------- */

  /* Pragmatic rather than RFC-complete: a stricter pattern rejects addresses
     that work, and the only authoritative test is delivery. */
  var cEmailPattern = /^[^\s@]+@[^\s@.]+(?:\.[^\s@.]+)+$/;

  function fnIsValidEmail(bv_vValue) {
    return cEmailPattern.test(fnTrim(bv_vValue));
  }

  /* ---------- Date of birth ---------------------------------------------- */

  function fnAgeInYears(bv_sBirth, bv_sReference) {
    var vAge = bv_sReference.getFullYear() - bv_sBirth.getFullYear();
    var vMonthGap = bv_sReference.getMonth() - bv_sBirth.getMonth();
    if (vMonthGap < 0 || (vMonthGap === 0 && bv_sReference.getDate() < bv_sBirth.getDate())) {
      vAge--;
    }
    return vAge;
  }

  /* Returns { vIsValid, vReason, vAge }.
     vReason is one of: ok | empty | malformed | future | underage | implausible
     Reasons are separated so the caller can say something useful rather than
     "invalid date". */
  function fnCheckBirthDate(bv_vISODate, bv_vMinAge, bv_vMaxAge) {
    var vValue = fnTrim(bv_vISODate);
    if (!vValue) { return { vIsValid: false, vReason: 'empty', vAge: null }; }

    var arParts = /^(\d{4})-(\d{2})-(\d{2})$/.exec(vValue);
    if (!arParts) { return { vIsValid: false, vReason: 'malformed', vAge: null }; }

    var vYear  = Number(arParts[1]);
    var vMonth = Number(arParts[2]);
    var vDay   = Number(arParts[3]);

    /* Built from parts, not parsed from the string: `new Date('2000-01-01')` is
       UTC midnight and shifts a day backwards in Brazilian time zones. */
    var sBirth = new Date(vYear, vMonth - 1, vDay);

    /* Date rolls 31/02 forward to 02/03 without complaining, so compare back. */
    if (sBirth.getFullYear() !== vYear ||
        sBirth.getMonth() !== vMonth - 1 ||
        sBirth.getDate() !== vDay) {
      return { vIsValid: false, vReason: 'malformed', vAge: null };
    }

    var sToday = new Date();
    sToday.setHours(0, 0, 0, 0);

    if (sBirth.getTime() > sToday.getTime()) {
      return { vIsValid: false, vReason: 'future', vAge: null };
    }

    var vAge = fnAgeInYears(sBirth, sToday);
    if (vAge > bv_vMaxAge) { return { vIsValid: false, vReason: 'implausible', vAge: vAge }; }
    if (vAge < bv_vMinAge) { return { vIsValid: false, vReason: 'underage', vAge: vAge }; }

    return { vIsValid: true, vReason: 'ok', vAge: vAge };
  }

  /* ISO date offset by a whole number of years, for the `min`/`max` attributes
     of a date input. Returns 'YYYY-MM-DD' in local time. */
  function fnISODateYearsAgo(bv_vYears) {
    var sDate = new Date();
    sDate.setFullYear(sDate.getFullYear() - bv_vYears);
    var vMonth = String(sDate.getMonth() + 1);
    var vDay   = String(sDate.getDate());
    return sDate.getFullYear() + '-' +
           (vMonth.length < 2 ? '0' + vMonth : vMonth) + '-' +
           (vDay.length   < 2 ? '0' + vDay   : vDay);
  }

  /* 'YYYY-MM-DD' -> 'DD/MM/YYYY'. String surgery on purpose: the value comes
     from a payload and must be shown exactly as sent, never re-derived. */
  function fnFormatISODateBR(bv_vISODate) {
    var arParts = /^(\d{4})-(\d{2})-(\d{2})/.exec(fnTrim(bv_vISODate));
    if (!arParts) { return fnTrim(bv_vISODate); }
    return arParts[3] + '/' + arParts[2] + '/' + arParts[1];
  }

  /* ---------- Whole numbers ---------------------------------------------- */

  /* Returns { vIsValid, vReason, vNumber }.
     vReason is one of: ok | empty | malformed | below | above
     'empty' is separate because on an optional field a blank is an answer. */
  function fnCheckInteger(bv_vValue, bv_vMin, bv_vMax) {
    var vValue = fnTrim(bv_vValue);
    if (!vValue) { return { vIsValid: false, vReason: 'empty', vNumber: null }; }

    /* Digits, optionally grouped with dots (cm_masks.fnThousandsBR). Not
       coerced: Number('1e3') is 1000 and Number('12,5') is NaN. */
    if (!/^\d{1,3}(?:\.\d{3})*$|^\d+$/.test(vValue)) {
      return { vIsValid: false, vReason: 'malformed', vNumber: null };
    }

    var vNumber = Number(vValue.replace(/\./g, ''));
    if (vNumber < bv_vMin) { return { vIsValid: false, vReason: 'below', vNumber: vNumber }; }
    if (vNumber > bv_vMax) { return { vIsValid: false, vReason: 'above', vNumber: vNumber }; }

    return { vIsValid: true, vReason: 'ok', vNumber: vNumber };
  }

  /* ---------- Website / Instagram ---------------------------------------- */

  function fnNormalizeWebsite(bv_vValue) {
    var vValue = fnTrim(bv_vValue);
    if (!vValue) { return ''; }
    if (!/^https?:\/\//i.test(vValue)) { vValue = 'https://' + vValue; }
    return vValue;
  }

  function fnIsValidWebsite(bv_vValue) {
    var vNormalized = fnNormalizeWebsite(bv_vValue);
    if (!vNormalized) { return false; }
    try {
      var sURL = new URL(vNormalized);
      return /^[a-z0-9-]+(?:\.[a-z0-9-]+)+$/i.test(sURL.hostname);
    } catch (sError) {
      return false;
    }
  }

  /* Accepts a handle, an @handle, or any form of profile URL, and returns the
     canonical '@handle'. People paste all three. */
  function fnNormalizeInstagram(bv_vValue) {
    var vValue = fnTrim(bv_vValue);
    if (!vValue) { return ''; }

    vValue = vValue.replace(/^https?:\/\//i, '')
                   .replace(/^www\./i, '')
                   .replace(/^(?:instagram\.com|instagr\.am)\//i, '');
    vValue = vValue.split(/[?#/]/)[0];
    vValue = vValue.replace(/^@+/, '');

    return vValue ? '@' + vValue : '';
  }

  function fnIsValidInstagram(bv_vValue) {
    var vHandle = fnNormalizeInstagram(bv_vValue).slice(1);
    return /^[A-Za-z0-9._]{1,30}$/.test(vHandle);
  }

  /* ---------- Files ------------------------------------------------------ */

  /* Returns { vIsValid, vReason } with vReason in: ok | empty | type | size */
  function fnCheckFile(bv_sFile, bv_arAcceptedTypes, bv_vMaxBytes) {
    if (!bv_sFile) { return { vIsValid: false, vReason: 'empty' }; }

    /* Some Android browsers hand over an empty type for a valid file, so the
       extension is the fallback rather than an outright rejection. */
    var vType = String(bv_sFile.type || '').toLowerCase();
    var vIsAccepted = bv_arAcceptedTypes.indexOf(vType) !== -1;

    if (!vIsAccepted && !vType) {
      var vName = String(bv_sFile.name || '').toLowerCase();
      vIsAccepted = bv_arAcceptedTypes.some(function (bv_vMime) {
        var vExtension = bv_vMime.split('/')[1];
        if (vExtension === 'jpeg') { return /\.jpe?g$/.test(vName); }
        return new RegExp('\\.' + vExtension + '$').test(vName);
      });
    }

    if (!vIsAccepted) { return { vIsValid: false, vReason: 'type' }; }
    if (bv_sFile.size > bv_vMaxBytes) { return { vIsValid: false, vReason: 'size' }; }

    return { vIsValid: true, vReason: 'ok' };
  }

  function fnFormatBytes(bv_vBytes) {
    if (bv_vBytes >= 1024 * 1024) {
      return (bv_vBytes / (1024 * 1024)).toFixed(bv_vBytes % (1024 * 1024) === 0 ? 0 : 1)
             .replace('.', ',') + ' MB';
    }
    return Math.ceil(bv_vBytes / 1024) + ' KB';
  }

  global.cm_validators = {
    fnOnlyDigits: fnOnlyDigits,
    fnTrim: fnTrim,
    fnOnlyCNPJChars: fnOnlyCNPJChars,
    fnCNPJCharValue: fnCNPJCharValue,
    fnCNPJCheckDigit: fnCNPJCheckDigit,
    fnIsValidCNPJ: fnIsValidCNPJ,
    fnCPFCheckDigit: fnCPFCheckDigit,
    fnIsValidCPF: fnIsValidCPF,
    fnIsValidRG: fnIsValidRG,
    fnIsValidPhoneBR: fnIsValidPhoneBR,
    fnIsMobilePhoneBR: fnIsMobilePhoneBR,
    fnIsValidEmail: fnIsValidEmail,
    fnCheckBirthDate: fnCheckBirthDate,
    fnAgeInYears: fnAgeInYears,
    fnISODateYearsAgo: fnISODateYearsAgo,
    fnFormatISODateBR: fnFormatISODateBR,
    fnCheckInteger: fnCheckInteger,
    fnNormalizeWebsite: fnNormalizeWebsite,
    fnIsValidWebsite: fnIsValidWebsite,
    fnNormalizeInstagram: fnNormalizeInstagram,
    fnIsValidInstagram: fnIsValidInstagram,
    fnCheckFile: fnCheckFile,
    fnFormatBytes: fnFormatBytes
  };

}(window));


/* -------------
   Usage Example
   -------------

cm_validators.fnIsValidCNPJ('12.ABC.345/01DE-35');   // true: masked or not, any case

*/
