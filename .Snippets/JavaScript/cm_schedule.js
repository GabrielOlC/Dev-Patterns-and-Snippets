/* -------------------------------------------------------------------------
   Source: FEBRACA Forms, page happyhour-isp-pet (wf_happyhour.js)
   Name: cm_schedule
   Purpose: Tells whether a form is open, from an optional opening and closing
       moment, and how long is left; runs a countdown to the next moment, with
       a sentence a screen reader can speak. The controller decides what to do
       about it.
   Context: Kept an event's sign-up closed outside its window and showed the
       countdown to its deadline.
   Requirements:
       - None. cm_countdown draws a clock that fnStartCountdown can drive: its
         dtClock is a valid bv_dtTargets.

   Parameters:
       - bv_vOpensAt / bv_vClosesAt: ISO 8601 with an explicit offset, always:
         '2026-10-08T16:00:00-03:00', not '2026-10-08 16:00'. Without the
         offset some engines read local time and others UTC, and the deadline
         moves by hours. Either may be empty.
       - bv_vNow: a fixed "now" for testing; production passes nothing.
       - bv_dtTargets (fnStartCountdown): the elements the caller owns, each
         optional; sText gets the spoken sentence:
           { sDays, sHours, sMinutes, sSeconds, sText }
       - bv_vsOnExpire: fires once, when the countdown reaches zero with the
         page open.
       - bv_dtCopy: optional, and so is each key; a key left out keeps the
         default, which speaks of a sign-up ("inscrições"). {vDays} and
         {vMoment} are filled in:
           {
             vOpensIn:    'Este formulário abre em {vDays}, em {vMoment}.',
             vClosedAt:   'As inscrições encerraram em {vMoment}.',
             vLastHours:  'Últimas horas: as inscrições encerram em {vMoment}.',
             vOpenFor:    'Este formulário fica disponível por {vDays}, até {vMoment}.',
             vToOpen:     'para a abertura das inscrições.',
             vToClose:    'para o encerramento das inscrições.',
             vOpeningNow: 'As inscrições estão abrindo agora.',
             vEnded:      'O prazo de inscrição terminou.'
           }
   Notes: fnCheckWindow's vState is one of:
       - always_open: no dates declared; nothing to show
       - before_open: the opening is in the future
       - open: inside the window, or open-ended
       - closed: the closing has passed
       fnDisplayState adds 'last_day', for display only. The clock is the
       visitor's, and a device clock can be moved: this explains the deadline
       to a late arrival; the real gate is server-side. An unparseable date is
       ignored (with an error on the console) rather than closing the form for
       everyone. The rest of each sentence is Brazilian Portuguese ("Faltam 2
       dias e 3 horas", "30/11/2026 às 23h59"). fnStartCountdown returns
       { vsStop }, or null when there is nothing to count.

   Version: 261030
   ------------------------------------------------------------------------- */

(function (global) {
  'use strict';

  var cOneDayMs = 24 * 60 * 60 * 1000;

  var dtDefaultCopy = {
    vOpensIn:    'Este formulário abre em {vDays}, em {vMoment}.',
    vClosedAt:   'As inscrições encerraram em {vMoment}.',
    vLastHours:  'Últimas horas: as inscrições encerram em {vMoment}.',
    vOpenFor:    'Este formulário fica disponível por {vDays}, até {vMoment}.',
    vToOpen:     'para a abertura das inscrições.',
    vToClose:    'para o encerramento das inscrições.',
    vOpeningNow: 'As inscrições estão abrindo agora.',
    vEnded:      'O prazo de inscrição terminou.'
  };

  /* The caller's words over the defaults, key by key. */
  function fnCopy(bv_dtCopy) {
    var dtCopy = {};
    for (var vKey in dtDefaultCopy) {
      if (Object.prototype.hasOwnProperty.call(dtDefaultCopy, vKey)) {
        dtCopy[vKey] = bv_dtCopy && bv_dtCopy[vKey] ? bv_dtCopy[vKey] : dtDefaultCopy[vKey];
      }
    }
    return dtCopy;
  }

  /* An unknown placeholder stays as written, so a typo shows on screen instead
     of vanishing. */
  function fnFill(bv_vText, bv_dtValues) {
    return bv_vText.replace(/\{(\w+)\}/g, function (bv_vMatch, bv_vKey) {
      return Object.prototype.hasOwnProperty.call(bv_dtValues, bv_vKey)
        ? bv_dtValues[bv_vKey]
        : bv_vMatch;
    });
  }

  function fnPad2(bv_vNumber) {
    return bv_vNumber < 10 ? '0' + bv_vNumber : String(bv_vNumber);
  }

  /* Whole days, rounded up, because "faltam 0 dias" reads as "it is over" when
     there are still eleven hours left. The last partial day counts as a day. */
  function fnDaysUntil(bv_sTarget, bv_sNow) {
    return Math.ceil((bv_sTarget.getTime() - bv_sNow.getTime()) / cOneDayMs);
  }

  /* null for an absent or unparseable date: a typo in the config reads as "no
     limit" (and an error on the console) instead of closing the form for all. */
  function fnParse(bv_vISO, bv_vLabel) {
    if (!bv_vISO) { return null; }

    var sDate = new Date(bv_vISO);
    if (isNaN(sDate.getTime())) {
      if (global.console && global.console.error) {
        global.console.error('cm_schedule: ' + bv_vLabel + ' is not a valid ISO date: ' +
          bv_vISO + '. The limit was ignored.');
      }
      return null;
    }
    return sDate;
  }

  /* Returns { vState, vDaysToOpen, vDaysToClose, sOpensAt, sClosesAt }. */
  function fnCheckWindow(bv_vOpensAt, bv_vClosesAt, bv_vNow) {
    var sNow    = bv_vNow ? new Date(bv_vNow) : new Date();
    var sOpens  = fnParse(bv_vOpensAt, 'bv_vOpensAt');
    var sCloses = fnParse(bv_vClosesAt, 'bv_vClosesAt');

    var dtResult = {
      vState: 'open',
      vDaysToOpen: null,
      vDaysToClose: null,
      sOpensAt: sOpens,
      sClosesAt: sCloses
    };

    if (!sOpens && !sCloses) {
      dtResult.vState = 'always_open';
      return dtResult;
    }

    if (sCloses && sNow.getTime() >= sCloses.getTime()) {
      dtResult.vState = 'closed';
      return dtResult;
    }

    if (sOpens && sNow.getTime() < sOpens.getTime()) {
      dtResult.vState = 'before_open';
      dtResult.vDaysToOpen = fnDaysUntil(sOpens, sNow);
      return dtResult;
    }

    if (sCloses) {
      dtResult.vDaysToClose = fnDaysUntil(sCloses, sNow);
    }
    return dtResult;
  }

  /* dd/mm/aaaa às HHhMM, in the reader's own locale offset. */
  function fnFormatMoment(bv_sDate) {
    if (!bv_sDate) { return ''; }

    return fnPad2(bv_sDate.getDate()) + '/' + fnPad2(bv_sDate.getMonth() + 1) + '/' +
           bv_sDate.getFullYear() + ' às ' + fnPad2(bv_sDate.getHours()) + 'h' +
           fnPad2(bv_sDate.getMinutes());
  }

  function fnPluralDays(bv_vDays) {
    return bv_vDays === 1 ? '1 dia' : bv_vDays + ' dias';
  }

  /* The sentence for the banner; '' means leave the banner hidden. */
  function fnDescribe(bv_dtWindow, bv_dtCopy) {
    var dtCopy = fnCopy(bv_dtCopy);

    if (bv_dtWindow.vState === 'always_open') { return ''; }

    if (bv_dtWindow.vState === 'before_open') {
      return fnFill(dtCopy.vOpensIn, {
        vDays: fnPluralDays(bv_dtWindow.vDaysToOpen),
        vMoment: fnFormatMoment(bv_dtWindow.sOpensAt)
      });
    }

    if (bv_dtWindow.vState === 'closed') {
      return fnFill(dtCopy.vClosedAt, { vMoment: fnFormatMoment(bv_dtWindow.sClosesAt) });
    }

    if (bv_dtWindow.vDaysToClose === null) { return ''; }

    /* Under a day the count is useless — "1 dia" for four hours left is worse
       than saying the hour. */
    if (bv_dtWindow.vDaysToClose <= 1) {
      return fnFill(dtCopy.vLastHours, { vMoment: fnFormatMoment(bv_dtWindow.sClosesAt) });
    }

    return fnFill(dtCopy.vOpenFor, {
      vDays: fnPluralDays(bv_dtWindow.vDaysToClose),
      vMoment: fnFormatMoment(bv_dtWindow.sClosesAt)
    });
  }

  /* The last day is a visual state only, so vState keeps four words. */
  function fnDisplayState(bv_dtWindow) {
    if (bv_dtWindow.vState === 'open' && bv_dtWindow.vDaysToClose !== null &&
        bv_dtWindow.vDaysToClose <= 1) {
      return 'last_day';
    }
    return bv_dtWindow.vState;
  }

  /* For a form that wants a line of prose instead of a clock. The element is
     the caller's; the service never goes looking for one. */
  function vsRenderBanner(bv_sElement, bv_dtWindow, bv_dtCopy) {
    if (!bv_sElement) { return; }

    var vText = fnDescribe(bv_dtWindow, bv_dtCopy);
    if (!vText) {
      bv_sElement.hidden = true;
      return;
    }

    bv_sElement.textContent = vText;
    bv_sElement.setAttribute('data-state', fnDisplayState(bv_dtWindow));
    bv_sElement.hidden = false;
  }

  /* ---------- Countdown ---------------------------------------------------
     The clock, as opposed to the sentence. Same window, read to the second.
     ------------------------------------------------------------------- */

  /* Days floored here, unlike fnDaysUntil: the clock also shows the hours. */
  function fnBreakdown(bv_sTarget, bv_vNow) {
    var sNow  = bv_vNow ? new Date(bv_vNow) : new Date();
    var vMs   = Math.max(0, bv_sTarget.getTime() - sNow.getTime());
    var vSecs = Math.floor(vMs / 1000);

    return {
      vTotalMs: vMs,
      vDays:    Math.floor(vSecs / 86400),
      vHours:   Math.floor((vSecs % 86400) / 3600),
      vMinutes: Math.floor((vSecs % 3600) / 60),
      vSeconds: vSecs % 60
    };
  }

  /* Which moment a clock counts to, or null when there is nothing to count:
     no window declared, or the window already closed. */
  function fnCountdownTarget(bv_dtWindow) {
    if (bv_dtWindow.vState === 'before_open') { return bv_dtWindow.sOpensAt; }
    if (bv_dtWindow.vState === 'open')        { return bv_dtWindow.sClosesAt; }
    return null;
  }

  function fnPluralUnit(bv_vValue, bv_vSingular, bv_vPlural) {
    return bv_vValue + ' ' + (bv_vValue === 1 ? bv_vSingular : bv_vPlural);
  }

  /* The digits as a sentence, for screen readers — built from the breakdown so
     both say the same. */
  function fnDescribeBreakdown(bv_dtLeft, bv_vState, bv_dtCopy) {
    var dtCopy = fnCopy(bv_dtCopy);
    var vIsEarly = bv_vState === 'before_open';

    if (bv_dtLeft.vTotalMs <= 0) {
      return vIsEarly ? dtCopy.vOpeningNow : dtCopy.vEnded;
    }

    var arParts   = [];
    var vLastUnit = 0;

    function vsPush(bv_vValue, bv_vSingular, bv_vPlural) {
      if (!bv_vValue) { return; }
      arParts.push(fnPluralUnit(bv_vValue, bv_vSingular, bv_vPlural));
      vLastUnit = bv_vValue;
    }

    vsPush(bv_dtLeft.vDays, 'dia', 'dias');
    vsPush(bv_dtLeft.vHours, 'hora', 'horas');
    vsPush(bv_dtLeft.vMinutes, 'minuto', 'minutos');

    /* Seconds only when they are all that is left. */
    if (!arParts.length) {
      vsPush(bv_dtLeft.vSeconds || 1, 'segundo', 'segundos');
    }

    var vList = arParts.length > 1
      ? arParts.slice(0, -1).join(', ') + ' e ' + arParts[arParts.length - 1]
      : arParts[0];

    /* Agreement: "Falta 1 dia", but "Faltam 4 dias" and "Faltam 1 hora e 1
       minuto" — a compound subject is plural however small its parts are. */
    var vVerb = (arParts.length === 1 && vLastUnit === 1) ? 'Falta ' : 'Faltam ';

    return vVerb + vList + ' ' + (vIsEarly ? dtCopy.vToOpen : dtCopy.vToClose);
  }

  /* Starts the clock. The digits are NOT a live region: announcing them every
     second would make the page unusable with a screen reader. */
  function fnStartCountdown(bv_dtTargets, bv_dtWindow, bv_vsOnExpire, bv_dtCopy) {
    var sTarget = fnCountdownTarget(bv_dtWindow);
    if (!sTarget) { return null; }

    var dtTargets   = bv_dtTargets || {};
    var dtCopy      = fnCopy(bv_dtCopy);
    var vTimer      = null;
    var vHasExpired = false;

    function vsWrite(bv_dtLeft) {
      if (dtTargets.sDays)    { dtTargets.sDays.textContent    = fnPad2(bv_dtLeft.vDays); }
      if (dtTargets.sHours)   { dtTargets.sHours.textContent   = fnPad2(bv_dtLeft.vHours); }
      if (dtTargets.sMinutes) { dtTargets.sMinutes.textContent = fnPad2(bv_dtLeft.vMinutes); }
      if (dtTargets.sSeconds) { dtTargets.sSeconds.textContent = fnPad2(bv_dtLeft.vSeconds); }
      if (dtTargets.sText) {
        dtTargets.sText.textContent = fnDescribeBreakdown(bv_dtLeft, bv_dtWindow.vState, dtCopy);
      }
    }

    function vsTick() {
      var dtLeft = fnBreakdown(sTarget);
      vsWrite(dtLeft);

      if (dtLeft.vTotalMs > 0 || vHasExpired) { return; }

      vHasExpired = true;
      if (vTimer !== null) { global.clearInterval(vTimer); }
      if (bv_vsOnExpire) { bv_vsOnExpire(); }
    }

    vsTick();
    if (!vHasExpired) { vTimer = global.setInterval(vsTick, 1000); }

    return {
      vsStop: function () {
        if (vTimer !== null) { global.clearInterval(vTimer); }
      }
    };
  }

  global.cm_schedule = {
    fnCheckWindow: fnCheckWindow,
    fnDescribe: fnDescribe,
    fnDisplayState: fnDisplayState,
    fnFormatMoment: fnFormatMoment,
    fnDaysUntil: fnDaysUntil,
    fnBreakdown: fnBreakdown,
    fnCountdownTarget: fnCountdownTarget,
    fnDescribeBreakdown: fnDescribeBreakdown,
    vsRenderBanner: vsRenderBanner,
    fnStartCountdown: fnStartCountdown
  };

}(window));


/* -------------
   Usage Example
   -------------

var dtWindow = cm_schedule.fnCheckWindow('2026-11-01T09:00:00-03:00', '2026-11-30T23:59:00-03:00');
if (dtWindow.vState === 'closed' || dtWindow.vState === 'before_open') { vsCloseForm(); }

*/
