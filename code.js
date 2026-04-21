(function (window) {
  "use strict";
  var MAX_PREVIEW_SWATCHES = 5;
  var SHOW_IN_HISTORY = false;
  var EXECUTE_IMMEDIATELY = true;

  var palettes = [
    {
      id: "ocean",
      name: "Ocean Blue",
      colors: ["#0A2463", "#247BA0", "#70C1B3", "#B2DBBF", "#F3FFBD", "#1D3557", "#F1FAEE", "#E63946", "#A8DADC", "#457B9D", "#1D3557", "#F1FAEE"]
    },
    {
      id: "sunset",
      name: "Sunset Warm",
      colors: ["#3D348B", "#7678ED", "#F7B801", "#F18701", "#F35B04", "#5F0F40", "#9A031E", "#FB8B24", "#E36414", "#0F4C5C", "#2D3142", "#FFFFFF"]
    },
    {
      id: "forest",
      name: "Forest Green",
      colors: ["#132A13", "#31572C", "#4F772D", "#90A955", "#ECF39E", "#355070", "#6D597A", "#B56576", "#E56B6F", "#EAAC8B", "#1B4332", "#FFFFFF"]
    },
    {
      id: "mono",
      name: "Modern Monochrome",
      colors: ["#111111", "#2E2E2E", "#5C5C5C", "#8A8A8A", "#CFCFCF", "#EAEAEA", "#F5F5F5", "#FFFFFF", "#3B82F6", "#0EA5E9", "#111111", "#FFFFFF"]
    }
  ];

  function renderPalettes() {
    var container = document.getElementById("paletteList");
    container.innerHTML = "";

    palettes.forEach(function (palette) {
      var button = document.createElement("button");
      button.className = "palette-btn";
      button.type = "button";
      button.textContent = palette.name;
      button.addEventListener("click", function () {
        applyPalette(palette);
      });

      var swatches = document.createElement("div");
      swatches.className = "swatches";

      palette.colors.slice(0, MAX_PREVIEW_SWATCHES).forEach(function (color) {
        var swatch = document.createElement("span");
        swatch.className = "swatch";
        swatch.style.backgroundColor = color;
        swatches.appendChild(swatch);
      });

      button.appendChild(swatches);
      container.appendChild(button);
    });
  }

  function setStatus(message, isError) {
    var status = document.getElementById("status");
    status.textContent = message;
    status.className = isError ? "error" : "";
  }

  function makeThemeScript(palette) {
    var colors = JSON.stringify(palette.colors);
    var themeName = JSON.stringify("Brand Palette - " + palette.name);

    return "(function(){" +
      "if(!Api || !Api.CreateTheme || !Api.GetPresentation){return {ok:false,error:'Presentation theme API unavailable. Please use a compatible ONLYOFFICE Presentation editor.'};}" +
      "var colors=" + colors + ";" +
      "var theme = Api.CreateTheme(" + themeName + ", colors);" +
      "Api.GetPresentation().SetTheme(theme);" +
      "return {ok:true};" +
      "})();";
  }

  function applyPalette(palette) {
    if (!window.Asc || !window.Asc.plugin || typeof window.Asc.plugin.callCommand !== "function") {
      setStatus("ONLYOFFICE plugin API is unavailable. Open this plugin from ONLYOFFICE Presentation editor.", true);
      return;
    }

    setStatus("Applying \"" + palette.name + "\"...");

    try {
      window.Asc.plugin.callCommand(
        makeThemeScript(palette),
        SHOW_IN_HISTORY,
        EXECUTE_IMMEDIATELY,
        function (result) {
          if (result && result.ok === false) {
            setStatus(result.error || "Failed to apply selected palette.", true);
            return;
          }
          setStatus("Applied \"" + palette.name + "\".");
        }
      );
    } catch (error) {
      setStatus(error && error.message ? error.message : "Failed to apply selected palette.", true);
    }
  }

  window.Asc = window.Asc || {};
  window.Asc.plugin = window.Asc.plugin || {};

  window.Asc.plugin.init = function () {
    renderPalettes();
    setStatus("Ready.");
  };

  window.Asc.plugin.button = function () {
    window.Asc.plugin.executeCommand("close", "");
  };
})(window);
