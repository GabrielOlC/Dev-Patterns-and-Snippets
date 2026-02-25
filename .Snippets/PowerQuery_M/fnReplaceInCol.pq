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
If you want the replacement to be exact WORD matched, then add an extra item to 
the list, as this:
    {"ColName1", { {"", "FindVal","ReplaceVal"}, {"", "FindVal.","ReplaceVal"} }},
 }
Note: Empty cells will become null; picked cols will all be lowered
------------------------------------------------------------------------- */
(bvTable as table, bvColumnReplaceValList as list) as table =>
    let
        //Set the list behaviour 
        fnReplaceVals = (bvTxt as text, bvReplaceWith as list) =>
            let 
                vReplacingVals = List.Accumulate(
                    bvReplaceWith,
                    Text.Lower(bvTxt),
                    (lmTxt, lmReplaceWith) => 
                        // Check if replacement is special (maybe one day change to a choose a type of replacement through picking a numb as it is on excel)
                        if List.Count(lmReplaceWith) = 3 then
                            // Break down each word based on space, and compare value
                            let
                                vWord = Text.Split(lmTxt, " "),
                                vReplacedWords = List.Transform(vWord, (lmWord) =>
                                    if lmWord = Text.Lower(lmReplaceWith{1}) then lmReplaceWith{2} else lmWord
                                )
                            in
                                Text.Combine(vReplacedWords, " ")
                        else
                        // Normal replacement
                        Text.Replace(lmTxt, Text.Lower( lmReplaceWith{0} ), lmReplaceWith{1})
                ),
                vNormalized = Text.Replace(vReplacingVals, " ", " ") //Double space is here
            in
                vNormalized,
            
        //Translate the list format of COL and Replace Vals list for fnReplaceVals
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
        {"CEP",{ {"00000000", ""} }}
    },

vCleanColumns = fnReplaceInColumn(
    fnRemoveColAccents(#"Removed Other Columns",List.Transform(vColumnReplaceValList, each _{0})), //could be just the #"Removed Other Columns"
    vColumnReplaceValList
)

*/