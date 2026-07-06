function escapeRegExp(string) { return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'); }

function cleanFontName(name) {
    var cleaned = name;
    cleaned = cleaned.replace(/^FONTSPRING DEMO\s*-\s*/i, '');
    cleaned = cleaned.replace(/^FONTSPRING DEMO\s*/i, '');
    cleaned = cleaned.replace(/^Free\s*-\s*/i, '');
    return cleaned;
}

function isFontBanned(dbName, aiFamily) {
    if (aiFamily.toLowerCase() === dbName.toLowerCase()) return true;
    var regexStrict = new RegExp("^" + escapeRegExp(dbName) + "(?:\\s|-|$)", "i");
    if (regexStrict.test(aiFamily)) return true;

    var dbZipped = dbName.replace(/\s+/g, '').toLowerCase();
    var aiZipped = aiFamily.replace(/\s+/g, ''); 
    if (dbZipped.length > 3 && aiZipped.toLowerCase().indexOf(dbZipped) === 0) {
        if (dbZipped.length === aiZipped.length) return true;
        var nextChar = aiZipped.charAt(dbZipped.length);
        if ((nextChar.toUpperCase() === nextChar && nextChar.toLowerCase() !== nextChar) || /[^a-zA-Z]/.test(nextChar)) {
            return true;
        }
    }
    return false;
}

function getIllustratorColor(hex, docSpace) {
    var r = parseInt(hex.substring(0, 2), 16);
    var g = parseInt(hex.substring(2, 4), 16);
    var b = parseInt(hex.substring(4, 6), 16);

    if (docSpace === DocumentColorSpace.CMYK) {
        var cmyk = new CMYKColor();
        var c = 1 - (r / 255); var m = 1 - (g / 255); var y = 1 - (b / 255);
        var k = Math.min(c, Math.min(m, y));
        if (k === 1) { cmyk.cyan = 0; cmyk.magenta = 0; cmyk.yellow = 0; cmyk.black = 100; } 
        else {
            cmyk.cyan = Math.round(((c - k) / (1 - k)) * 100);
            cmyk.magenta = Math.round(((m - k) / (1 - k)) * 100);
            cmyk.yellow = Math.round(((y - k) / (1 - k)) * 100);
            cmyk.black = Math.round(k * 100);
        }
        return cmyk;
    } else {
        var rgb = new RGBColor();
        rgb.red = r; rgb.green = g; rgb.blue = b;
        return rgb;
    }
}

function processDocument(jsonPayload) {
    if (app.documents.length === 0) return "WARN_NODOC";
    var doc = app.activeDocument;
    var textFrames = doc.textFrames;
    if (textFrames.length === 0) return "WARN_NOTEXT";

    var payload;
    try { payload = eval('(' + jsonPayload + ')'); } catch(e) { return "ERR_FORMAT"; }

    var bannedFonts = payload.fonts;
    var actions = payload.actions;
    if (!bannedFonts || bannedFonts.length === 0) return "ERR_EMPTYDB";

    var targetColor = null;
    if (actions.recolor && !actions.del) targetColor = getIllustratorColor(actions.hex, doc.documentColorSpace);

    var uniqueDocFonts = {};
    for (var i = 0; i < textFrames.length; i++) {
        try {
            if (textFrames[i].textRange.characters.length === 0) continue;
            var fam = cleanFontName(textFrames[i].textRange.characters[0].characterAttributes.textFont.family);
            uniqueDocFonts[fam] = true;
        } catch(e) {}
    }

    var bannedInDoc = {};
    var hasBanned = false;
    var foundFontsArray = [];

    for (var docFont in uniqueDocFonts) {
        for (var j = 0; j < bannedFonts.length; j++) {
            if (isFontBanned(bannedFonts[j], docFont)) {
                bannedInDoc[docFont] = true;
                hasBanned = true;
                foundFontsArray.push(docFont);
                break;
            }
        }
    }

    if (!hasBanned) return "CLEAN";

    for (var k = textFrames.length - 1; k >= 0; k--) {
        var frame = textFrames[k];
        try {
            if (frame.textRange.characters.length === 0) continue;
            var currentFam = cleanFontName(frame.textRange.characters[0].characterAttributes.textFont.family);
            
            if (bannedInDoc[currentFam]) {
                if (actions.del) { frame.remove(); continue; }
                if (actions.select) frame.selected = true;
                if (actions.strike) frame.textRange.characterAttributes.strikeThrough = true;
                if (actions.recolor && targetColor !== null) {
                    frame.textRange.characterAttributes.fillColor = targetColor;
                    if (frame.textRange.characterAttributes.strokeColor && frame.textRange.characterAttributes.strokeColor.typename !== "NoColor") {
                        frame.textRange.characterAttributes.strokeColor = targetColor;
                    }
                }
            }
        } catch(err) {}
    }

    var jsonResult = "[";
    for (var f = 0; f < foundFontsArray.length; f++) {
        jsonResult += '"' + foundFontsArray[f].replace(/"/g, '\\"') + '"';
        if (f < foundFontsArray.length - 1) jsonResult += ",";
    }
    jsonResult += "]";
    
    return jsonResult;
}