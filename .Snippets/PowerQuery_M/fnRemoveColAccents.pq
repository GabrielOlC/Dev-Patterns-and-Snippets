/* -------------------------------------------------------------------------
   Source: AMPARA CCPA Data Analysis qr_Geocoding
   Function: fnRemoveColAccents
   Purpose: Replaces accented characters (Latin-1) with their ASCII equivalents
			for the whole requested column
   Context: Critical for standardizing names/addresses before performing Database
            Merges or ID generation.
   Requirement: fnRemoveStrAccents
   ------------------------------------------------------------------------- */
(bvTable as table, bvCols as list) as table =>
    let
        // Runs for each column
        vColumnReplaceList = List.Transform(
            bvCols,
            (lmCol) => { lmCol, each fnRemoveStrAccents(_), type text }
        ),

        // Apply transformations
        v = Table.TransformColumns(
            bvTable,
            vColumnReplaceList
        )
    in
        v