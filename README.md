# A Simple Interactive Metro Route Planner

Built entirely in Excel and then in Access—without using Dijkstra’s algorithm.

## Overview  
The project uses Lignes.txt as its sole data source. That file contains three columns—line, time, and station—and, for each direction of every metro line, lists the stations served along with the access time (in seconds) from the originating terminus. From this single table, routes are generated without ever building an explicit graph or implementing Dijkstra’s algorithm: all connections (direct, one-change or two-change) are produced via relational joins between intermediate tables (Lignes, Corresp, etc.), with only a fixed set of combinations (0, 1 or 2 transfers).

In Excel, Power Query carries out these joins and builds the final Trajets table in one Refresh—pre-warmed in Workbook_Open—so the user experiences virtually no latency on the first button click. In Access, equivalent SQL queries following the very same logic are compiled and optimized on demand by the Jet/ACE engine, and routes refresh automatically whenever the departure or arrival station is changed.

## Quick Start  
 
> ⚠️ **Macro Security**  
> Office blocks macros by default—always review the code or use an isolated sandbox/VM before enabling.

1. **Download**  
   - `Demo.xlsm` (Excel) or `Demo.accdb` (Access)

2. **Open**  
   - In Excel/Access, go to **File → Open** (macros stay disabled)

3. **Enable**  
   - Click **Enable Content** **only** after you’ve reviewed the code or are running in an offline sandbox/VM

4. **Use**  
   - **Excel**: Select your **Departure** and **Arrival**, then click the button   
   - **Access**: Select **Departure** and **Arrival**—the results update automatically
## Contributing  
Feel free to submit issues or pull requests for enhancements or bug fixes.
