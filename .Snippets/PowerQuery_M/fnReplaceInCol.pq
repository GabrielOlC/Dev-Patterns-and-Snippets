/* -------------------------------------------------------------------------
    Source: AMPARA CCPA Data Analysis qr_Geocoding
    Function: fnReplaceInCol
    Purpose: Replace values into a table across multiple columns and multiple
            "find" -> "Replace to" values. Also removes double spaces and trim
    Context: Used to normalize the Address Strings Components before aggregating it
            to use the Geocoding API'
    bvColumnReplaceValList Expected Structure:
    {
        {"ColName1", { {"FindVal","ReplaceVal"}, {"FindVal.","ReplaceVal"} }},
        {"ColName2", { {"FindVal","ReplaceVal"}, {"FindVal.","ReplaceVal"} }},
        ...
    }

    Replacement Rules:
        {"FindVal","ReplaceVal"}
            - Normal replacement (Text.Replace). Any match inside the cell is replaced.

        {"Mode","FindVal","ReplaceVal"}
            - Allows special replacement behaviour depending on Mode:

            Mode = "" | 0 | "byWord"
                Break down each word based on spaces and compare each token.
                Replacement happens only when the token exactly matches FindVal.

            Mode = 1 | "byCell"
                Compare the entire cell value (trimmed and lowered) against FindVal.
                Replacement only happens when the whole cell matches exactly.
            
            mode = 2 | "byLen<"
                Compare the cell len against FindVal. 
                Replacement happens only when cell len is smaller or equal to Findval.

    Examples:
        { {"ColName1", { {"byWord","rd","road"} }}, 
          {"ColName1", { {"byLen","1",""} }},
          ...
        }

    Note: Empty cells will become null; selected cols will all be lowered
    ------------------------------------------------------------------------- */
    (bvTable as table, bvColumnReplaceValList as list) as table =>
        let
            //Set the inner list behaviour { {"FindVal","ReplaceVal"}, {"FindVal.","ReplaceVal"} }}
            fnReplaceVals = (bvText as text, bvReplaceWith as list) =>
                let 
                    vReplacingVals = List.Accumulate(
                        bvReplaceWith,
                        if bvText = null then "" else Text.Trim(Text.Lower(bvText)),
                        (lmTxt, lmReplaceWith) =>
                            // Check if replacement is special (maybe one day change to a choose a type of replacement through picking a numb as it is on excel)
                            if List.Count(lmReplaceWith) = 3 then
                                let
                                    // normalize mode and find/replace values (work in lowercase for matching)
                                    vMode = Text.Lower( Text.From(lmReplaceWith{0}) ),
                                    vFind = Text.Lower( Text.From(lmReplaceWith{1}) ),
                                    vReplace = Text.From(lmReplaceWith{2})
                                in
                                    // Whole-word match: Break down each word based on space, and compare value before replacing
                                    if vMode = "" or vMode = "0" or vMode = "byword" then
                                        let
                                            vWords = Text.Split(lmTxt, " "),
                                            vReplacedWords = List.Transform(vWords, (lmWord) => 
                                                if lmWord = vFind then vReplace else lmWord
                                            )
                                        in
                                            Text.Combine(vReplacedWords, " ")
                                            
                                    // whole-cell match: replace only when the entire trimmed cell equals the find value		
                                    else if vMode = "1" or vMode = "bycell" then
                                        if lmTxt = vFind then vReplace else lmTxt
                                    
                                    // length match
                                    else if vMode = "2" or vMode = "bylen<" then
                                        if Text.Length(lmTxt) < Number.FromText(vFind) then vReplace else lmTxt

                                    // fallback: Do nothing
                                    else
                                        lmTxt
                            else
                                // Normal replacement
                                Text.Replace(lmTxt, Text.Lower(lmReplaceWith{0}), lmReplaceWith{1})
                    ),
                    vNormalized = Text.Replace(vReplacingVals, "  ", " ") //Double space is here
                in
                    vNormalized,
                
            //Translate the list format of COL and Replace Vals list for fnReplaceVals. That is, we remove the col from the original list and give the replace list to fnReplaceVals
            vColumnReplaceList = List.Transform(
                bvColumnReplaceValList,
                (lmColRep) => { lmColRep{0}, each Text.Trim(fnReplaceVals(_, lmColRep{1})), type text } //TRIM is here
            ),

            vTransformCols = Table.TransformColumns(
                bvTable,
                vColumnReplaceList
            ),

            //Replace empty cells with null for each modified column
			v = Table.ReplaceValue(vTransformCols, "", null, Replacer.ReplaceValue, List.Transform(bvColumnReplaceValList, each _{0}))

        in
            v
			
/* --- USAGE EXAMPLE ---

vColumnReplaceValList = 
    { 
        {"Logradouro", { 
            {"Rua", ""}, {"", "R.",""}, {"", "R", ""}, {"", "Eua", ""}, 

            {"Avenida Dos", ""}, {"Av. Dos",""}, {"Av dos",""},
            {"Avenida Das", ""}, {"Av. Das",""}, {"Av das",""},
            {"Avenida Do", ""}, {"Av. Do",""}, {"Av do",""},
            {"Avenida Da", ""}, {"Av. Da",""}, {"Av da",""},
            {"Avenida", ""}, {"", "Av.",""}, {"", "Av",""}, 

            {"Estrada do", ""}, {"Estrada da", ""}, 
            {"Estrada dos", ""}, {"Estrada das", ""},
            {"Estrada", ""},

            {"Travessa", ""}, {"", "Tv.", ""}, {"", "Tv", ""}

        }},
        {"DA", { {"Selecione",""} }},
        {"CEP",{ {"00000000", ""}, {"bylen<", "8", ""} }}
    },

vCleanColumns = fnReplaceInColumn(
    fnRemoveColAccents(#"Removed Other Columns",List.Transform(vColumnReplaceValList, each _{0})), //could be just the #"Removed Other Columns"
    vColumnReplaceValList
)

*/