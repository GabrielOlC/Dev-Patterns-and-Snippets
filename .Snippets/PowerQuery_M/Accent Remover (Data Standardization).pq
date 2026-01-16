/* -------------------------------------------------------------------------
   Function: fnRemoveAccents
   Purpose: Replaces accented characters (Latin-1) with their ASCII equivalents.
   Context: Critical for standardizing names/addresses before performing Database
            Merges or ID generation.
   ------------------------------------------------------------------------- */

(fnText as nullable text) as text =>
let
    vTtxt = if fnText = null then "" else fnText,
    vPairs = {
        {"á","a"},{"à","a"},{"ã","a"},{"â","a"},{"ä","a"},
        {"Á","A"},{"À","A"},{"Ã","A"},{"Â","A"},{"Ä","A"},
        {"é","e"},{"è","e"},{"ê","e"},{"ë","e"},
        {"É","E"},{"È","E"},{"Ê","E"},{"Ë","E"},
        {"í","i"},{"ì","i"},{"î","i"},{"ï","i"},
        {"Í","I"},{"Ì","I"},{"Î","I"},{"Ï","I"},
        {"ó","o"},{"ò","o"},{"õ","o"},{"ô","o"},{"ö","o"},
        {"Ó","O"},{"Ò","O"},{"Õ","O"},{"Ô","O"},{"Ö","O"},
        {"ú","u"},{"ù","u"},{"û","u"},{"ü","u"},
        {"Ú","U"},{"Ù","U"},{"Û","U"},{"Ü","U"},
        {"ç","c"},{"Ç","C"},{"ñ","n"},{"Ñ","N"},
        {"ý","y"},{"ỳ","y"},{"ÿ","y"},{"Ý","Y"}
    },
    result = List.Accumulate(vPairs, vTtxt , (state, pair) => Text.Replace(state, pair{0}, pair{1}))
in
    result
