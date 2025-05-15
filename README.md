# Resource Scheduling Solution with Microsoft Excel

This guide will walk you through implementing a resource scheduling system to tie together your project engagements and personnel availability datasets using Excel.

## Setup Instructions

### 1. Create Data Tables
First, set up four worksheets with these tables:

#### Projects Sheet
| ProjectID | ProjectName | ClientName | StartDate | EndDate | Status | PriorityLevel |
|-----------|------------|------------|-----------|---------|--------|--------------|
| P001 | Website Redesign | Acme Corp | 5/20/2025 | 8/15/2025 | Active | High |
| P002 | Mobile App Development | TechStart Inc | 6/1/2025 | 9/30/2025 | Active | Medium |
| ... | ... | ... | ... | ... | ... | ... |

#### Requirements Sheet
| RequirementID | ProjectID | RoleNeeded | SkillsRequired | HoursPerWeek | StartDate | EndDate |
|--------------|-----------|------------|---------------|--------------|-----------|---------|
| R001 | P001 | Project Manager | Project Management/Client Communication | 20 | 5/20/2025 | 8/15/2025 |
| R002 | P001 | UI Designer | UI/UX Design/Wireframing | 30 | 5/25/2025 | 7/10/2025 |
| ... | ... | ... | ... | ... | ... | ... |

#### Personnel Sheet
| EmployeeID | EmployeeName | Role | Skills | MaxHoursPerWeek | CostRate | Location |
|------------|--------------|------|--------|----------------|----------|----------|
| E001 | John Smith | Project Manager | Project Management/Client Communication/Agile | 40 | 125 | HQ |
| E002 | Emily Johnson | Project Manager | Project Management/Mobile/Scrum Master | 40 | 130 | HQ |
| ... | ... | ... | ... | ... | ... | ... |

#### Assignments Sheet
| AssignmentID | EmployeeID | ProjectID | RequirementID | StartDate | EndDate | HoursPerWeek | Status |
|--------------|------------|-----------|--------------|-----------|---------|--------------|--------|
| A001 | E001 | P001 | R001 | 5/20/2025 | 8/15/2025 | 20 | Confirmed |
| A002 | E003 | P001 | R002 | 5/25/2025 | 7/10/2025 | 30 | Confirmed |
| ... | ... | ... | ... | ... | ... | ... | ... |

### 2. Format as Tables
- Select each data range and press Ctrl+T to format as tables
- Give each table a name (Projects, Requirements, Personnel, Assignments)

### 3. Create Resource Availability Sheet

1. Create a new worksheet called "Availability"
2. Set up a table with these columns:
   - EmployeeID
   - EmployeeName
   - Role
   - MaxHoursPerWeek
   - AssignedHours
   - AvailableHours

3. Use these formulas:
   - EmployeeID, EmployeeName, Role, MaxHoursPerWeek: Pull from Personnel table
   - AssignedHours: `=SUMIFS(Assignments[HoursPerWeek], Assignments[EmployeeID], [@EmployeeID], Assignments[Status], "Confirmed")`
   - AvailableHours: `=[@MaxHoursPerWeek]-[@AssignedHours]`

### 4. Create Resource Allocation Matrix

1. Create a new worksheet called "Allocation"
2. Create a PivotTable:
   - Select the Assignments table
   - Insert > PivotTable
   - Place fields:
     - Rows: EmployeeID, EmployeeName
     - Columns: Create calculated dates by week
     - Values: Sum of HoursPerWeek
   - Add conditional formatting to show workload intensity

### 5. Create Resource Matching Tool

1. Create a new worksheet called "Matching"
2. Create a table with these columns:
   - RequirementID
   - ProjectID
   - ProjectName
   - RoleNeeded
   - SkillsRequired
   - HoursPerWeek
   - StartDate
   - EndDate
   - MatchingEmployees
   - RecommendedMatch

3. Use these formulas:
   - RequirementID, ProjectID, RoleNeeded, etc.: Pull from Requirements table
   - ProjectName: `=VLOOKUP([@ProjectID], Projects, 2, FALSE)`
   - MatchingEmployees: Use a complex FILTER formula (see below)
   - RecommendedMatch: First available matching employee

### 6. Create Dashboard

1. Create a dashboard sheet with:
   - Resource utilization chart
   - Project timeline view
   - Resource gap indicators
   - Slicers for filtering

## Key Formulas

### Resource Matching Formula
```
=TEXTJOIN(", ", TRUE, 
  IF(COUNTIFS(Personnel[Skills], "*"&[@SkillsRequired]&"*", 
             AvailabilityTable[AvailableHours], ">="&[@HoursPerWeek]), 
      FILTER(Personnel[EmployeeName], 
             (ISNUMBER(SEARCH([@SkillsRequired], Personnel[Skills])))*(AvailabilityTable[AvailableHours]>=[@HoursPerWeek])),
      "No match"))
```

### Resource Allocation by Week
```
=SUMIFS(Assignments[HoursPerWeek], 
        Assignments[EmployeeID], [@EmployeeID], 
        Assignments[StartDate], "<="&WeekEndDate, 
        Assignments[EndDate], ">="&WeekStartDate)
```

### Resource Gap Calculation
```
=IF(SUMIFS(Assignments[HoursPerWeek], Assignments[RequirementID], [@RequirementID])<[@HoursPerWeek],
  "UNDERSTAFFED: "&TEXT([@HoursPerWeek]-SUMIFS(Assignments[HoursPerWeek], Assignments[RequirementID], [@RequirementID]), "0")&" hours needed",
  "Fully staffed")
```

## Advanced Features

### Adding Data Validation
Set up data validation for key fields:
- Status: Create dropdown lists
- Priority: High, Medium, Low, Critical
- Skills: Use a separate skills table

### Setting Up Weekly Calendar View
1. Create a worksheet with weeks as columns
2. List employees as rows
3. Use conditional formatting to show assignments

### Automating with VBA
Simple refresh button:
```vba
Sub RefreshResourceData()
    Sheets("Availability").Calculate
    Sheets("Allocation").PivotTables("PivotAllocation").RefreshTable
    Sheets("Dashboard").Calculate
End Sub
```

## Tips for Maintaining the System

1. **Regular Updates**: Update assignments weekly
2. **Version Control**: Save dated copies of your file
3. **Data Validation**: Use dropdown lists to prevent errors
4. **Documentation**: Document formulas and calculations
5. **Backup**: Create regular backups of your scheduling file
