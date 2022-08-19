How long to continue to keep data exporting after student has left (needs to be for followups or allowing exit process to complete correctly) - use days as value

KCY Could be static (loaded from cachefile)
KCY Could be literal (just simply dump the description)
KCY Could be padded (pad the year numbers)

Add OliverFields variable/flag to exporter - In addition to normal fields it requires the following

Pull in KCY file
    Only do if set to literal or padded, otherwise pull master copy

Pull in Student Records
    Copy out of CSV into Working Table all ACTV,LVNG and FUT
    Run through all LEFT records, comparing exit date to current date - extra timeframe
        Add to Working copy if in-date
    Process email and username from working copy based upon settings

Pull in Staff Records
    Copy out of CSV into Working Table all ACTV
    Run through all LEFT records comparing exit date to current date - extra timeframe
        Add to Working copy if in-date
    Process email and username from working copy based upon settings

Pull in DF (Family) records
    Process Family records into working table only if the DFKey exists in SF.FAMILY
    IF ST.CONTACT_A = B
    Put B into A fields

Pull in UM (Address) records
    Process Address records into working table only if the UMKEY exists in DF.HOMEKEY
    Process Address records into working table only if the UMKEY exists in SF.HOMEKEY