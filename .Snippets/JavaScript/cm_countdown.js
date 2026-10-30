/* -------------------------------------------------------------------------
   Source: FEBRACA Forms, page happyhour-isp-pet (the clock's markup and
       base.css, driven by wf_happyhour.js)
   Name: cm_countdown
   Purpose: Draws a countdown clock (days, hours, minutes and seconds in boxes,
       a label above, a caption below and a sentence for screen readers) and
       hands back its parts, so any ticker can drive it.
   Context: The sign-up clock of an event page: navy, with the digits turning
       orange on the last day. It was page markup and page CSS before it
       became a service.
   Technique: The styles travel inside this file and are attached once, as a
       constructable stylesheet where the browser has one, since a
       Content-Security-Policy without 'unsafe-inline' blocks a <style>
       element but not that; elsewhere, as a <style> element. Selectors are
       one class each, so a page's element resets (p { margin: 0 }) don't
       undo them. Every colour is a custom property with a fallback; anything
       else is overridden with a selector stronger than one class.
   Requirements:
       - None to draw. To tick: cm_schedule.fnStartCountdown takes dtClock as
         its bv_dtTargets; any other source calls dtClock.vsShow.

   Parameters:
       - bv_sHost: the element whose content the clock replaces; it is
         unhidden.
       - bv_dtOptions: optional, and so is each key:
           {
             vLabel:       'Tempo restante para se inscrever',
             vCaption:     'Até 30/11/2026 às 23h59',
             arUnits:      ['vDays', 'vHours', 'vMinutes', 'vSeconds'],
             dtUnitLabels: { vDays: 'dias', vHours: 'horas', vMinutes: 'min', vSeconds: 'seg' },
             vIsUrgent:    false
           }
         arUnits picks which boxes show, and in which order; a unit label left
         out keeps its default. vIsUrgent turns the digits to the accent.
   Notes: fnRender returns dtClock, or null without a host:
           { sRoot, sDays, sHours, sMinutes, sSeconds, sText, vsShow, vsSetUrgent }
       The s* are the elements, null for a unit left out: the shape
       cm_schedule.fnStartCountdown writes into. vsShow(bv_dtLeft, bv_vSpoken)
       writes a breakdown { vDays, vHours, vMinutes, vSeconds } and, when
       given, the sentence. The digits are aria-hidden and never a live
       region, since announcing them every second would make the page
       unusable with a screen reader; sText is read when the person reaches
       it. Colours, each falling back to the project's token, then to the
       FEBRACA value the clock was drawn with:
           --countdown-bg          → --brand-navy   → #212B5E
           --countdown-text        → --brand-white  → #FFFFFF
           --countdown-accent      → --brand-orange → #F79A22
           --countdown-unit-bg     → rgba(255, 255, 255, 0.08)
           --countdown-unit-border → rgba(255, 255, 255, 0.18)
       Set them in the page's CSS, or with
       sHost.style.setProperty('--countdown-bg', '#0B0D14'), which a CSP
       allows where a style attribute is blocked.

   Version: 261030
   ------------------------------------------------------------------------- */

(function (global) {
  'use strict';

  var arDefaultUnits = ['vDays', 'vHours', 'vMinutes', 'vSeconds'];

  var dtDefaultUnitLabels = {
    vDays: 'dias',
    vHours: 'horas',
    vMinutes: 'min',
    vSeconds: 'seg'
  };

  /* Where each unit's element goes in dtClock: the names cm_schedule writes into. */
  var dtUnitTargets = {
    vDays: 'sDays',
    vHours: 'sHours',
    vMinutes: 'sMinutes',
    vSeconds: 'sSeconds'
  };

  var cPlaceholder = '--';

  var cCSS = [
    '.countdown {',
    '  background: var(--countdown-bg, var(--brand-navy, #212B5E));',
    '  color: var(--countdown-text, var(--brand-white, #FFFFFF));',
    '  border-radius: var(--radius, 8px);',
    '  padding: var(--sp-4, 1rem);',
    '  text-align: center;',
    '}',
    '.countdown__label {',
    '  margin: 0;',
    '  font-size: var(--fs-100, 0.8125rem);',
    '  text-transform: uppercase;',
    '  letter-spacing: 0.08em;',
    '  opacity: 0.85;',
    '}',
    '.countdown__units {',
    '  display: flex;',
    '  justify-content: center;',
    '  gap: var(--sp-2, 0.5rem);',
    '  margin: var(--sp-3, 0.75rem) 0 0;',
    '  padding: 0;',
    '  list-style: none;',
    '}',
    /* White at low alpha over the surface, not a new colour. */
    '.countdown__unit {',
    '  flex: 1 1 0;',
    '  max-width: 5.5rem;',
    '  background: var(--countdown-unit-bg, rgba(255, 255, 255, 0.08));',
    '  border: 1px solid var(--countdown-unit-border, rgba(255, 255, 255, 0.18));',
    '  border-radius: var(--radius-sm, 4px);',
    '  padding: var(--sp-2, 0.5rem) var(--sp-1, 0.25rem);',
    '}',
    /* tabular-nums, or the row shifts every second as digit widths change. */
    '.countdown__value {',
    '  display: block;',
    '  font-size: 1.75rem;',
    '  font-weight: 700;',
    '  line-height: 1.1;',
    '  font-variant-numeric: tabular-nums;',
    '}',
    '.countdown__unit-label {',
    '  display: block;',
    '  margin-top: var(--sp-1, 0.25rem);',
    '  font-size: var(--fs-100, 0.8125rem);',
    '  text-transform: uppercase;',
    '  letter-spacing: 0.06em;',
    '  opacity: 0.8;',
    '}',
    '.countdown__caption {',
    '  margin: var(--sp-3, 0.75rem) 0 0;',
    '  font-size: var(--fs-200, 0.9375rem);',
    '  opacity: 0.85;',
    '}',
    '.countdown--urgent .countdown__value {',
    '  color: var(--countdown-accent, var(--brand-orange, #F79A22));',
    '}',
    '.countdown__spoken {',
    '  position: absolute;',
    '  width: 1px;',
    '  height: 1px;',
    '  margin: -1px;',
    '  padding: 0;',
    '  overflow: hidden;',
    '  clip: rect(0 0 0 0);',
    '  clip-path: inset(50%);',
    '  white-space: nowrap;',
    '  border: 0;',
    '}',
    '@media (min-width: 40em) {',
    '  .countdown { padding: var(--sp-5, 1.5rem); }',
    '  .countdown__value { font-size: 2.25rem; }',
    '}'
  ].join('\n');

  var vHasStyles = false;

  function vsAttachStyles() {
    if (vHasStyles) { return; }
    vHasStyles = true;

    if (global.CSSStyleSheet && 'adoptedStyleSheets' in document) {
      try {
        var sSheet = new global.CSSStyleSheet();
        sSheet.replaceSync(cCSS);
        document.adoptedStyleSheets = document.adoptedStyleSheets.concat([sSheet]);
        return;
      } catch (sError) {
        /* A browser with the property but not the constructor: the <style>
           element below. */
      }
    }

    var sStyle = document.createElement('style');
    sStyle.textContent = cCSS;
    document.head.appendChild(sStyle);
  }

  function fnElement(bv_vTag, bv_vClass, bv_vText) {
    var sElement = document.createElement(bv_vTag);
    sElement.className = bv_vClass;
    if (bv_vText) { sElement.textContent = bv_vText; }
    return sElement;
  }

  function fnPad2(bv_vNumber) {
    return bv_vNumber < 10 ? '0' + bv_vNumber : String(bv_vNumber);
  }

  function fnRender(bv_sHost, bv_dtOptions) {
    if (!bv_sHost) { return null; }

    var dtOptions = bv_dtOptions || {};
    var arUnits   = dtOptions.arUnits && dtOptions.arUnits.length ? dtOptions.arUnits : arDefaultUnits;
    var dtLabels  = dtOptions.dtUnitLabels || {};

    vsAttachStyles();

    var sRoot   = fnElement('div', 'countdown');
    var dtClock = {
      sRoot: sRoot, sDays: null, sHours: null, sMinutes: null, sSeconds: null, sText: null
    };

    if (dtOptions.vLabel) { sRoot.appendChild(fnElement('p', 'countdown__label', dtOptions.vLabel)); }

    var sUnits = fnElement('ul', 'countdown__units');
    sUnits.setAttribute('aria-hidden', 'true');

    arUnits.forEach(function (bv_vUnit) {
      if (!dtUnitTargets[bv_vUnit]) {
        if (global.console && global.console.error) {
          global.console.error('cm_countdown: unknown unit ' + bv_vUnit + ', left out.');
        }
        return;
      }

      var sUnit  = fnElement('li', 'countdown__unit');
      var sValue = fnElement('span', 'countdown__value', cPlaceholder);
      sUnit.appendChild(sValue);
      sUnit.appendChild(fnElement('span', 'countdown__unit-label',
        dtLabels[bv_vUnit] || dtDefaultUnitLabels[bv_vUnit]));
      sUnits.appendChild(sUnit);
      dtClock[dtUnitTargets[bv_vUnit]] = sValue;
    });

    sRoot.appendChild(sUnits);

    dtClock.sText = fnElement('p', 'countdown__spoken');
    sRoot.appendChild(dtClock.sText);

    if (dtOptions.vCaption) { sRoot.appendChild(fnElement('p', 'countdown__caption', dtOptions.vCaption)); }

    dtClock.vsShow = function (bv_dtLeft, bv_vSpoken) {
      for (var vUnit in dtUnitTargets) {
        var sValue = dtClock[dtUnitTargets[vUnit]];
        if (sValue && typeof bv_dtLeft[vUnit] === 'number') {
          sValue.textContent = fnPad2(bv_dtLeft[vUnit]);
        }
      }
      if (typeof bv_vSpoken === 'string') { dtClock.sText.textContent = bv_vSpoken; }
    };

    dtClock.vsSetUrgent = function (bv_vIsUrgent) {
      sRoot.classList.toggle('countdown--urgent', !!bv_vIsUrgent);
    };

    dtClock.vsSetUrgent(dtOptions.vIsUrgent);

    while (bv_sHost.firstChild) { bv_sHost.removeChild(bv_sHost.firstChild); }
    bv_sHost.appendChild(sRoot);
    bv_sHost.hidden = false;

    return dtClock;
  }

  global.cm_countdown = {
    fnRender: fnRender
  };

}(window));


/* -------------
   Usage Example
   -------------

var dtWindow = cm_schedule.fnCheckWindow('', '2026-11-30T23:59:00-03:00');
var dtClock  = cm_countdown.fnRender(document.getElementById('countdown'), {
  vLabel:    'Tempo restante para se inscrever',
  vCaption:  'Até ' + cm_schedule.fnFormatMoment(cm_schedule.fnCountdownTarget(dtWindow)),
  vIsUrgent: cm_schedule.fnDisplayState(dtWindow) === 'last_day'
});
cm_schedule.fnStartCountdown(dtClock, dtWindow, function () { location.reload(); });

*/
