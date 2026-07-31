* Clear and set up
clear all
set more off

* Define file path and years
* Project root = the folder that CONTAINS this /code directory.
global root "C:\Users\Hugo Morão\Desktop\dgeg_fuel_organized"
global datapath "$root\data\processed\dataset_fuel_sales.xlsx"
global years "1995 1996 1997 1998 1999 2000 2001 2002 2003 2004 2005 2006 2007 2008 2009 2010 2011 2012 2013 2014 2015 2016 2017 2018 2019 2020 2021 2022 2023 2024 2025 2026"

* Initialize empty dataset - DON'T save empty observations
tempfile master
clear
gen year = .
gen month = .
gen product = ""
gen consumption = .
save `master', replace empty  // Use empty option

* Loop through years
foreach year of global years {
    
    * Import data from year tab (import as strings to preserve names)
    import excel "$datapath", sheet("`year'") firstrow allstring clear
    
    * First column is month (A variable after firstrow import)
    rename A month_name
    
    * Drop total row and empty rows
    drop if month_name == "TOTAL" | missing(month_name)
    
    * Generate month number
    gen month = .
    replace month = 1 if month_name == "Janeiro"
    replace month = 2 if month_name == "Fevereiro"
    replace month = 3 if month_name == "Março"
    replace month = 4 if month_name == "Abril"
    replace month = 5 if month_name == "Maio"
    replace month = 6 if month_name == "Junho"
    replace month = 7 if month_name == "Julho"
    replace month = 8 if month_name == "Agosto"
    replace month = 9 if month_name == "Setembro"
    replace month = 10 if month_name == "Outubro"
    replace month = 11 if month_name == "Novembro"
    replace month = 12 if month_name == "Dezembro"
    destring month, replace
    
    * Drop if month is missing
    drop if missing(month)
    
    * Get product variables
    ds month_name month, not
    local product_vars `r(varlist)'
    
    * Store original variable labels (which contain full names)
    tempfile year_data
    clear
    gen year = .
    gen month = .
    gen product = ""
    gen consumption = .
    save `year_data', replace empty  // Use empty option
    
    * Reload data
    import excel "$datapath", sheet("`year'") firstrow allstring clear
    rename A month_name
    drop if month_name == "TOTAL" | missing(month_name)
    
    gen month = .
    replace month = 1 if month_name == "Janeiro"
    replace month = 2 if month_name == "Fevereiro"
    replace month = 3 if month_name == "Março"
    replace month = 4 if month_name == "Abril"
    replace month = 5 if month_name == "Maio"
    replace month = 6 if month_name == "Junho"
    replace month = 7 if month_name == "Julho"
    replace month = 8 if month_name == "Agosto"
    replace month = 9 if month_name == "Setembro"
    replace month = 10 if month_name == "Outubro"
    replace month = 11 if month_name == "Novembro"
    replace month = 12 if month_name == "Dezembro"
    destring month, replace
    
    drop if missing(month)
    
    * Get product variables again
    ds month_name month, not
    local product_vars `r(varlist)'
    
    * Process each product
    foreach prod of local product_vars {
        preserve
        
        * Get full product name from variable label or use varname
        local prod_label : variable label `prod'
        if "`prod_label'" == "" {
            local prod_name "`prod'"
        }
        else {
            local prod_name "`prod_label'"
        }
        
        keep month `prod'
        gen product = "`prod_name'"
        rename `prod' consumption_str
        
        * Clean consumption
        replace consumption_str = subinstr(consumption_str, "*", "", .)
        replace consumption_str = subinstr(consumption_str, " ", "", .)
        replace consumption_str = "" if consumption_str == "."
        
        gen consumption = real(consumption_str)
        gen year = `year'
        keep year month product consumption
        
        * Drop if all key variables are missing
        drop if missing(year) | missing(month) | missing(product)
        
        * Append to year data
        append using `year_data'
        save `year_data', replace
        restore
    }
    
    * Append year data to master
    use `year_data', clear
    drop if missing(year) | missing(month) | missing(product)
    append using `master'
    save `master', replace
}

* Final dataset preparation
use `master', clear
drop if missing(year) | missing(month) | missing(product)

* Check for duplicates before panel declaration
duplicates report product year month
duplicates tag product year month, gen(dup)
list product year month consumption if dup > 0, sepby(product)

* Remove duplicates (keep first occurrence)
duplicates drop product year month, force

drop dup

* Create time variable
gen date_monthly = ym(year, month)
format date_monthly %tm

* Declare panel structure
encode product, gen(product_id)
xtset product_id date_monthly

* Sort and order
sort product year month
order date_monthly product consumption, first

* Save final dataset
save "$root\output\oil_products_panel.dta", replace

* Summary
xtsum consumption
tab product
bysort product: summarize consumption


* Start fresh from the saved panel data
use "$root\output\oil_products_panel.dta", clear

preserve

* Keep only what we need
keep date_monthly product consumption

* Create product_id and store the mapping
encode product, gen(product_id)

* Create a local macro with all product mappings
quietly levelsof product_id, local(all_ids)
foreach id of local all_ids {
    quietly levelsof product if product_id == `id', local(pname) clean
    local prod_`id' "`pname'"
}

* Keep only reshape variables
keep date_monthly product_id consumption

* Reshape
reshape wide consumption, i(date_monthly) j(product_id)

* Apply labels using the stored mappings
foreach id of local all_ids {
    capture label variable consumption`id' "`prod_`id''"
}

* Sort by date
sort date_monthly
format date_monthly %tm

* Export to Excel (product labels as the header row)
export excel using "$root\output\oil_products_wide.xlsx", firstrow(varlabels) replace

* Export the same wide table to CSV with the SAME label headers.
* Note: Stata's `export delimited` writes variable NAMES, not labels, so it
* cannot reproduce the label-header format directly. The companion script
* code/xlsx_to_csv.py converts the workbook above into oil_products_wide.csv
* (UTF-8-BOM). Run it after this do-file, or via: python code/xlsx_to_csv.py

restore

