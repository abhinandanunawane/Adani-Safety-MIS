(function () {
  var inp = document.getElementById("kpi-doc-search");
  var status = document.getElementById("kpi-doc-search-status");
  if (!inp) return;

  function run() {
    var q = (inp.value || "").trim().toLowerCase();
    var sections = document.querySelectorAll(
      ".doc-page .doc-kpi-cat:not(.doc-kpi-cat--source)"
    );
    var vis = 0;
    sections.forEach(function (sec) {
      var rows = sec.querySelectorAll("tbody tr");
      var secVis = 0;
      rows.forEach(function (tr) {
        var show = !q || tr.textContent.toLowerCase().indexOf(q) !== -1;
        tr.hidden = !show;
        if (show) {
          secVis++;
          vis++;
        }
      });
      sec.hidden = q.length > 0 && secVis === 0;
    });
    if (status) {
      if (!q) {
        status.textContent = "";
      } else if (vis === 0) {
        status.textContent = "No matching KPIs.";
      } else {
        status.textContent = "Showing " + vis + " matching row" + (vis === 1 ? "" : "s") + ".";
      }
    }
  }

  inp.addEventListener("input", run);
  inp.addEventListener("search", run);
})();
