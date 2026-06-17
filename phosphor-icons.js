(function () {
  var css = [
    ".ph-bold{display:inline-flex;align-items:center;justify-content:center;width:1em;height:1em;font-style:normal;font-weight:700;line-height:1;vertical-align:-.125em}",
    ".ph-bold:before{content:'◆';font-size:.72em}",
    ".ph-activity:before{content:'↯'}",
    ".ph-users:before{content:'US'}",
    ".ph-ambulance:before{content:'+'}",
    ".ph-warning-circle:before,.ph-warning:before{content:'!'}",
    ".ph-heartbeat:before{content:'H'}",
    ".ph-calendar-blank:before{content:'□'}",
    ".ph-clock:before{content:'◷'}",
    ".ph-file-text:before{content:'≡'}",
    ".ph-caret-down:before{content:'⌄'}",
    ".ph-upload-simple:before{content:'↑'}",
    ".ph-file-csv:before{content:'CSV';font-size:.42em}",
    ".ph-arrow-counter-clockwise:before{content:'↺'}",
    ".ph-brain:before{content:'●'}",
    ".ph-car:before{content:'▰'}",
    ".ph-siren:before{content:'◉'}",
    ".ph-stethoscope:before{content:'S'}",
    ".ph-download-simple:before{content:'↓'}",
    ".ph-user-list:before{content:'☷'}",
    ".ph-trash:before{content:'×'}",
    ".ph-plus:before{content:'+'}"
  ].join("");
  var style = document.createElement("style");
  style.textContent = css;
  document.head.appendChild(style);
})();
