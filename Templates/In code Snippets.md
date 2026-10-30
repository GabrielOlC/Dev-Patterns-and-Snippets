### Template

Use the language's own comment syntax (`/* */` in M and JS, `'` in VBA, and so on); the fields are same.

```
/* -------------------------------------------------------------------------
   Source: Project and module where the code was first developed
   Name: Function / equation name* (how the client calls it; also the file name)
   Purpose: What does it do and why?*
   Context: Where was it created and how was it used? (ideally 1 line, up to 1 paragraph)
   Technique: How does it work? (optional, up to 1 paragraph, only when not obvious from the code)
   Requirements:
       - Which other services does it need to work? (linked)
       - Library installations go in the "Requirements list" at the end

   Parameters:
       - bvArgument: what it holds, with its expected structure as a generic example (non-obvious ones only)
   Modes: What each mode does (only if an argument switches behaviour)
   Notes: Side effects (e.g., values lowered, blanks turned to null)

   Attribution:
       - Original Concept: Where it came from (author, link, license — ideally linked)
       - Changes: What differs from the original (if applicable)

    Version: The date the code part of the service was created or last updated (YYMMDD)
   ------------------------------------------------------------------------- */

<> Paste the code here <>


/* -------------
   Usage Example
   -------------

<> Paste ONE example of how the client calls this service, with masked data <>

*/

/* -------------
   Requirements list (if applicable)
   -------------

<> For Python, R or other projects with a requirements.txt <>

*/
```

### Filled example

```powerquery
/* -------------------------------------------------------------------------
   Source: AMPARA CCPA Data Analysis, query qr_Geocoding
   Name: fnReplaceInCol
   Purpose: Replaces values across several columns using find-and-replace rules per column, then trims and removes double spaces.
   Context: Normalized address components before sending them to a geocoding API.
   Technique: Each column's rules are applied in sequence with List.Accumulate, so a rule sees the result of the one before it.
   Parameters:
       - bvColumnReplaceValList: one entry per column, each with its list of rules:
           {
             {"ColName1", { {"FindVal", "ReplaceVal"}, {"Mode", "FindVal", "ReplaceVal"} }},
             {"ColName2", { {"FindVal", "ReplaceVal"} }}
           }
         A 2-item rule replaces FindVal anywhere in the cell; a 3-item rule uses a mode.
   Modes:
       - "" | 0 | "byWord": replaces only words that exactly match FindVal
       - 1 | "byCell": replaces only when the whole cell matches FindVal
       - 2 | "byLen<": replaces the whole cell when its length is below FindVal
       - 3 | "byContains": replaces the whole cell when it contains FindVal
   Notes: Selected columns are lowered and trimmed; empty cells become null.

   Version: 20261030
   ------------------------------------------------------------------------- */

<> function code <>


/* -------------
   Usage Example
   -------------

vRules = {
    {"Street",     { {"Avenida", ""}, {"", "Av.", ""} }},
    {"PostalCode", { {"00000000", ""}, {"bylen<", "8", ""} }}
},
vClean = fnReplaceInCol(#"Previous Step", vRules)

*/
```
