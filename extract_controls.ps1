 = Get-Content 'C:\OSBusiness FE\FORMGEN.frm' -TotalCount 5152
 = @()
foreach ( in ) {
    if ( -match '^\s+Begin\s+(\S+)\s+(\S+)') {
         = [1]
         = [2]
         +=  	
    }
}
 | Sort-Object | Get-Unique
