/* -------------------------------------------------------------------------
   Source: FEBRACA Forms, page cadastro-afiliado (wf_afiliado.js)
   Name: cm_masks
   Purpose: Formats inputs as they are typed (CNPJ including the alphanumeric
       one, CPF, CEP, Brazilian phone, whole reais) and keeps the caret where
       the person expects it.
   Context: Masked the document, CEP, phone and amount fields of the FEBRACA
       forms, one vsBind per field.
   Technique: A pattern marks value slots with '0'; every other character is
       a separator. The value fills it and stops when it runs out, so no
       dangling separator shows. The binder places the caret by counting the
       characters that carry value, not every character, so moving separators
       never drag it. Backspacing onto a separator removes the neighbouring
       value character instead, since the mask would put the separator back.

   Parameters:
       - bv_fnFormat: one of the masks (fnCNPJ, fnCPF, fnCEP, fnPhoneBR,
         fnThousandsBR), or any function from raw text to formatted text.
       - bv_vContentPattern: a one-character RegExp for what the mask counts
         as value; digits when omitted. fnCNPJ needs cCNPJCharPattern, or the
         caret jumps once the value holds letters.
   Notes: Pure string work plus one binder; knows nothing about any form.
       vsBind formats a value already in the field when it binds. The CNPJ is
       uppercased, with letters only in the first twelve positions. fnPhoneBR
       picks the 10- or 11-digit pattern from how much has been typed.
       fnThousandsBR has no centavos and drops leading zeros.

   Version: 261030
   ------------------------------------------------------------------------- */

(function (global) {
  'use strict';

  /* '0' marks a value slot; every other character is a literal separator. */
  var cPatternCNPJ  = '00.000.000/0000-00';
  var cPatternCPF   = '000.000.000-00';
  var cPatternCEP   = '00000-000';
  var cPatternTel10 = '(00) 0000-0000';
  var cPatternTel11 = '(00) 00000-0000';

  var cMaxCharsCNPJ   = 14;
  var cMaxDigitsCPF   = 11;
  var cMaxDigitsCEP   = 8;
  var cMaxDigitsPhone = 11;
  var cMaxDigitsMoney = 12;   /* R$ 999.999.999.999 — far past any ONG's books */

  /* Only the first twelve positions of a CNPJ may hold letters; the two check
     digits are always numeric. */
  var cCNPJBaseLength = 12;

  /* What each mask counts as a character of value. Lowercase is included in the
     CNPJ pattern because the person types it before the mask uppercases it, and
     the caret is measured against what they typed. */
  var cDigitPattern    = /[0-9]/;
  var cCNPJCharPattern = /[0-9A-Za-z]/;

  function fnOnlyDigits(bv_vValue) {
    return String(bv_vValue === null || bv_vValue === undefined ? '' : bv_vValue)
      .replace(/\D+/g, '');
  }

  /* ---------- Content characters ----------------------------------------- */

  function fnCountContent(bv_vValue, bv_vPattern) {
    var vCount = 0;
    for (var z = 0; z < bv_vValue.length; z++) {
      if (bv_vPattern.test(bv_vValue.charAt(z))) { vCount++; }
    }
    return vCount;
  }

  function fnOnlyContent(bv_vValue, bv_vPattern) {
    var vOut = '';
    for (var z = 0; z < bv_vValue.length; z++) {
      var vChar = bv_vValue.charAt(z);
      if (bv_vPattern.test(vChar)) { vOut += vChar; }
    }
    return vOut;
  }

  /* Fills a pattern and stops as soon as the value runs out, so a half-typed
     entry never shows a dangling separator. */
  function fnApplyPattern(bv_vValue, bv_vPattern) {
    var vOut = '';
    var vIndex = 0;
    for (var z = 0; z < bv_vPattern.length && vIndex < bv_vValue.length; z++) {
      var vChar = bv_vPattern.charAt(z);
      if (vChar === '0') {
        vOut += bv_vValue.charAt(vIndex);
        vIndex++;
      } else {
        vOut += vChar;
      }
    }
    return vOut;
  }

  /* ---------- The masks --------------------------------------------------- */

  /* Uppercased: the check-digit arithmetic reads the ASCII code. A letter in
     the two check-digit positions never appears — those are numeric by rule. */
  function fnOnlyCNPJChars(bv_vValue) {
    var vChars = String(bv_vValue === null || bv_vValue === undefined ? '' : bv_vValue)
      .toUpperCase()
      .replace(/[^0-9A-Z]+/g, '')
      .slice(0, cMaxCharsCNPJ);

    return vChars.slice(0, cCNPJBaseLength) +
           vChars.slice(cCNPJBaseLength).replace(/[^0-9]+/g, '');
  }

  function fnCNPJ(bv_vValue) {
    return fnApplyPattern(fnOnlyCNPJChars(bv_vValue), cPatternCNPJ);
  }

  function fnCPF(bv_vValue) {
    return fnApplyPattern(fnOnlyDigits(bv_vValue).slice(0, cMaxDigitsCPF), cPatternCPF);
  }

  function fnCEP(bv_vValue) {
    return fnApplyPattern(fnOnlyDigits(bv_vValue).slice(0, cMaxDigitsCEP), cPatternCEP);
  }

  /* Landlines are ten digits and mobiles eleven, so the pattern is chosen from
     how much has been typed rather than fixed up front. */
  function fnPhoneBR(bv_vValue) {
    var vDigits = fnOnlyDigits(bv_vValue).slice(0, cMaxDigitsPhone);
    var vPattern = vDigits.length <= 10 ? cPatternTel10 : cPatternTel11;
    return fnApplyPattern(vDigits, vPattern);
  }

  /* Whole reais with thousand separators, no centavos: typing a comma is where
     many people get stuck, and padded decimals would throw the caret off.
     Leading zeros are dropped. */
  function fnThousandsBR(bv_vValue) {
    var vDigits = fnOnlyDigits(bv_vValue)
      .replace(/^0+(?=\d)/, '')
      .slice(0, cMaxDigitsMoney);
    return vDigits.replace(/\B(?=(\d{3})+$)/g, '.');
  }

  /* ---------- Caret ------------------------------------------------------- */

  /* Character offset that sits just after the Nth content character, then past
     any separators glued to it, so the next keystroke lands where it looks like
     it will. */
  function fnCaretAfterContent(bv_vFormatted, bv_vCount, bv_vPattern) {
    if (bv_vCount <= 0) { return 0; }

    var vSeen = 0;
    for (var z = 0; z < bv_vFormatted.length; z++) {
      if (!bv_vPattern.test(bv_vFormatted.charAt(z))) { continue; }
      vSeen++;
      if (vSeen !== bv_vCount) { continue; }

      var vPos = z + 1;
      while (vPos < bv_vFormatted.length && !bv_vPattern.test(bv_vFormatted.charAt(vPos))) { vPos++; }
      return vPos;
    }
    return bv_vFormatted.length;
  }

  /* Backspacing onto a separator deletes a character the mask immediately puts
     back, which reads as a dead key. Drop the neighbouring content character
     instead. */
  function fnReapplyDeletion(bv_vRawValue, bv_vCountBefore, bv_vIsForward, bv_fnFormat, bv_vPattern) {
    var arChars = fnOnlyContent(bv_vRawValue, bv_vPattern).split('');

    if (bv_vIsForward) {
      if (bv_vCountBefore >= arChars.length) { return null; }
      arChars.splice(bv_vCountBefore, 1);
      return { vFormatted: bv_fnFormat(arChars.join('')), vCountBefore: bv_vCountBefore };
    }

    if (bv_vCountBefore <= 0) { return null; }
    arChars.splice(bv_vCountBefore - 1, 1);
    return { vFormatted: bv_fnFormat(arChars.join('')), vCountBefore: bv_vCountBefore - 1 };
  }

  /* bv_vContentPattern says which characters this mask keeps, and defaults to
     digits. Pass cCNPJCharPattern along with fnCNPJ — without it the caret is
     counted in digits while the value holds letters, and it jumps. */
  function vsBind(bv_sInput, bv_fnFormat, bv_vContentPattern) {
    if (!bv_sInput) { return; }

    var vContentPattern = bv_vContentPattern || cDigitPattern;

    if (bv_sInput.value) { bv_sInput.value = bv_fnFormat(bv_sInput.value); }
    var vPrevious = bv_sInput.value;

    bv_sInput.addEventListener('input', function (bv_sEvent) {
      var vRawValue = bv_sInput.value;

      var vCaret = bv_sInput.selectionStart;
      if (vCaret === null || vCaret === undefined) { vCaret = vRawValue.length; }

      var vCountBefore = fnCountContent(vRawValue.slice(0, vCaret), vContentPattern);
      var vFormatted   = bv_fnFormat(vRawValue);

      var vInputType  = bv_sEvent && bv_sEvent.inputType ? bv_sEvent.inputType : '';
      var vIsDeleting = vInputType.indexOf('delete') === 0;

      if (vIsDeleting && vFormatted === vPrevious) {
        var dtRedo = fnReapplyDeletion(
          vRawValue, vCountBefore, vInputType === 'deleteContentForward',
          bv_fnFormat, vContentPattern);
        if (dtRedo) {
          vFormatted   = dtRedo.vFormatted;
          vCountBefore = dtRedo.vCountBefore;
        }
      }

      if (vFormatted !== vRawValue) {
        bv_sInput.value = vFormatted;
        var vNewCaret = fnCaretAfterContent(vFormatted, vCountBefore, vContentPattern);
        /* Some browsers refuse setSelectionRange on certain input types. The
           value is still correct, so a failure here is cosmetic. */
        try { bv_sInput.setSelectionRange(vNewCaret, vNewCaret); } catch (sError) { /* noop */ }
      }

      vPrevious = bv_sInput.value;
    });
  }

  global.cm_masks = {
    cDigitPattern: cDigitPattern,
    cCNPJCharPattern: cCNPJCharPattern,
    fnOnlyDigits: fnOnlyDigits,
    fnCountContent: fnCountContent,
    fnOnlyContent: fnOnlyContent,
    fnApplyPattern: fnApplyPattern,
    fnCaretAfterContent: fnCaretAfterContent,
    fnOnlyCNPJChars: fnOnlyCNPJChars,
    fnCNPJ: fnCNPJ,
    fnCPF: fnCPF,
    fnCEP: fnCEP,
    fnPhoneBR: fnPhoneBR,
    fnThousandsBR: fnThousandsBR,
    vsBind: vsBind
  };

}(window));


/* -------------
   Usage Example
   -------------

cm_masks.vsBind(document.getElementById('fmCNPJ'), cm_masks.fnCNPJ, cm_masks.cCNPJCharPattern);
// typing 12abc34501de35 shows 12.ABC.345/01DE-35

*/
