Clearly read Architecture.md file
Dont create a module unless specified

# Overview
App is developed based on Energy Institute Guidelines
T2.8 SURGE/MOMENTUM CHANGES DUE TO VALVE OPERATION

## Workflow
- user enter the [Valve Tag] and [CaseType] in the sheet ValveList in the table "tbValveList"
- Based on Valve  Tag and CaseType in the tbInput columns with the valve tag is generarated after column "Notes"


## Naming Convection
- [ColumnName] - string inside square bracket denotes the column name of the named table
- table name starts with "tb" or "ls" eg "tbValveList" or "lsCaseType"

## Sheets
- ValveList---contain tbValveList
- Data --- contain the tables for the data validation dropdown purpose
- Inputs -- where all the necessary input for the calc is entered by user
- Ref -- contains a table tbRequiredInput which describes the required input for the different CaseType
- Results -- after calculation results are summarized here in a printable format.

## Tables
- tbValveList columns Valve Tag,	CaseType
- lsCase columsn Case
- lsValveType col Valve Type
- lsSupportType cols Support Type,	Theta
- tbInput cols Parameter,	Symbol,	Unit,	Notes, and then list valves entered in tbValveList
- tbRequiredInput cols Parameter ,	Symbol, 	Units, 	Valve Closure, 	Valve Opening (Liquid/Multiphase),	Valve Opening (Dry Gas)
- tbResult cols Valve Tag,	CaseType,	Valve Type,	Support Type,	Fmax (kN),	Flim (kN),	LOF (-),	Flag

## tbInput

[Parameter]
Case Type
Fluid density
Ratio of Specific Heat Capacities (Cp/Cv)
Speed of sound
External Main Line Diameter
Internal Main Line Diameter
Young’s Modulus of main line material
Fluid Bulk Modulus
Upstream Pipe Length
Molecular Weight
Upstream Static Pressure
Pump head at zero flow
Vapour Pressure
Static Pressure drop
Universal Gas Constant
Main line Wall Thickness
Valve Closing Time
Upstream Temperature
Steady State Fluid Velocity
Mass Flow Rate
Pipe Support Type
Main line Wall Thickness for SCH 40
Valve Type

### UI/UX
when user click generate button in Input sheet, then columns are created after [Notes] in tbInput based on the valve tags.
Case type row is filled based on the case type selected in the tbValveList for the corresponding valve tag.
- based on the case type selected in the tbValveList, the required input parameters are highlighted in the tbInput for the corresponding valve tag column. Required input parameters are fetched from the tbRequiredInput table in the Ref sheet.
- If tbRequiredInput[Parameter][#CaseType]= 1 then that parameter is required for that case type if 0 then not required.
- Required input parameters are highlighted with light yellow color.
- Not Required input parameters are greyed out and locked for editing.
- User can enter the input parameters in the highlighted cells only.
- Similarly for the other valve tag columns in the tbInput table.

### VBA Calculation Logic
- Public Type ValveInputs contains all the input parameters from tbInput[Parameter]
- Public Type CalculationResult contains all the output parameters to be written in tbResult
- once all the input are gathered based on the case type, appropriate calculation function is called and results are stored in CalculationResult type and then written to tbResult table.
- Before calculation, input validation is performed to check if all the required input parameters are entered by user. If any required parameter is missing, an error message is shown to the user and calculation is aborted.

## tbRequiredInput
Parameter	Symbol	Units	Valve Closure	Valve Opening (Liquid/Multiphase)	Valve Opening (Dry Gas)
Fluid density	ρ	kg/m³	1	1	0
Ratio of Specific Heat Capacities (Cp/Cv)	γ	–	0	0	1
Speed of sound	c	m/s	1	0	0
External Main Line Diameter	Dext	mm	1	0	1
Internal Main Line Diameter	Dint	mm	1	0	1
Young’s Modulus of main line material	Eml	N/m²	1	0	0
Fluid Bulk Modulus	K	N/m²	1	0	0
Upstream Pipe Length	Lup	m	1	0	0
Molecular Weight	Mw	g/mol	0	0	1
Upstream Static Pressure	P1	Pa	1	1	0
Pump head at zero flow	Pshut-in	Pa	1	0	0
Vapour Pressure	Pv	Pa	0	0	0
Static Pressure drop	∆P	Pa	0	0	0
Universal Gas Constant	R	J/K·kmol	0	0	1
Main line Wall Thickness	T	mm	1	1	1
Valve Closing Time	Tclose	s	1	0	0
Upstream Temperature	Te	K	0	0	1
Steady State Fluid Velocity	v	m/s	1	0	0
Mass Flow Rate	W	kg/s	0	1	1
Pipe Support Type	–	–	1	1	1
Main line Wall Thickness for SCH 40	–	mm	1	1	1
Valve Type	–	–	1	0	0

## lsCase
Case
Valve Closure
Valve Opening (Liquid/Multiphase)
Valve Opening (Dry Gas)

## lsValveList
Valve Type
Ball-Full
Ball-Reduced
Butterfly
Globe
Gate

## lsSupportType
Support Type	Theta
Stiff	4
Medium Stiff	2
Medium	1
Flexible	0.2


## Units
Follow the units specified in the tbRequiredInput[Units]

## Calculation Background

all the 3 cases have same final formula ;
LOF = Fmax / Flim

where Flim has common formula for all the 3 cases

only Fmax formula is different for all the 3 cases



# Codebase Architecture
## Modules
1. CalculationEngine.bas
**Purpose**: All calculation logic and mathematical functions
2. DataStructures.bas
**Purpose**: Data types and validation
3. UIManager.bas
**Purpose**: UI interactions and worksheet management
4. DevTools.bas
**Purpose**: Development tools and utilities for Import and Export to and from vcode env to excel
### Worksheet Specific MOdules
For event handling of specific worksheets the naming convection is "Module_<WorksheetName>.bas"
1. Module_Inputs.bas
2. Module_Ref.bas
3. Module_Results.bas
4. Module_ValveList.bas
