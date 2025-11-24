# Business Continuity Plan - Copilot Studio Implementation
# Project Plan with Gantt Chart

Add-Type -AssemblyName System.Drawing

$doc = New-Object -ComObject Word.Application
$doc.Visible = $false
$document = $doc.Documents.Add()
$selection = $doc.Selection

# Title Page
$selection.Font.Size = 24
$selection.Font.Bold = $true
$selection.TypeText("Business Continuity Plan")
$selection.TypeParagraph()
$selection.TypeText("Copilot Studio Implementation")
$selection.TypeParagraph()
$selection.Font.Size = 16
$selection.Font.Bold = $false
$selection.TypeParagraph()
$selection.TypeText("Project Plan & Timeline")
$selection.TypeParagraph()
$selection.TypeParagraph()
$selection.Font.Size = 12
$selection.TypeText("Date: November 23, 2025")
$selection.TypeParagraph()
$selection.TypeText("Duration: 8-10 weeks")
$selection.TypeParagraph()
$selection.TypeText("Team Size: 3-4 people")
$selection.InsertBreak(7) # Page break

# Project Overview
$selection.Font.Size = 18
$selection.Font.Bold = $true
$selection.TypeText("Project Overview")
$selection.TypeParagraph()
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()
$selection.TypeText("Scope: Implementation of Copilot Studio solution for Business Continuity Plan document Q&A")
$selection.TypeParagraph()
$selection.TypeText("Documents: ~50 documents across multiple SharePoint sites with folder hierarchies")
$selection.TypeParagraph()
$selection.TypeText("Key Objectives:")
$selection.TypeParagraph()
$selection.TypeText("  • Enable natural language Q&A for BCP documents")
$selection.TypeParagraph()
$selection.TypeText("  • Maintain existing SharePoint security and permissions")
$selection.TypeParagraph()
$selection.TypeText("  • Provide fast, accurate responses with proper citations")
$selection.TypeParagraph()
$selection.TypeText("  • Ensure enterprise-grade compliance and governance")
$selection.TypeParagraph()
$selection.TypeParagraph()

# Team Structure
$selection.Font.Size = 14
$selection.Font.Bold = $true
$selection.TypeText("Team Structure")
$selection.TypeParagraph()
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()
$selection.TypeText("  • Copilot Studio Developer (1)")
$selection.TypeParagraph()
$selection.TypeText("  • SharePoint Administrator (1)")
$selection.TypeParagraph()
$selection.TypeText("  • BCP Subject Matter Expert (1)")
$selection.TypeParagraph()
$selection.TypeText("  • Project Manager (1)")
$selection.TypeParagraph()
$selection.TypeParagraph()

# Insert Gantt Chart section
$selection.InsertBreak(7) # Page break
$selection.Font.Size = 18
$selection.Font.Bold = $true
$selection.TypeText("Project Gantt Chart")
$selection.TypeParagraph()
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()

# Create Gantt Chart Table
$ganttTable = $selection.Tables.Add($selection.Range, 38, 13)
$ganttTable.Borders.Enable = $true
$ganttTable.Style = "Grid Table 4 - Accent 1"

# Headers
$ganttTable.Cell(1,1).Range.Text = "Phase"
$ganttTable.Cell(1,2).Range.Text = "Task"
$ganttTable.Cell(1,3).Range.Text = "Days"
for ($week = 1; $week -le 10; $week++) {
    $ganttTable.Cell(1, $week + 3).Range.Text = "W$week"
}

# Phase 1: Discovery & Planning
$row = 2
$ganttTable.Cell($row,1).Range.Text = "Phase 1"
$ganttTable.Cell($row,2).Range.Text = "SharePoint Audit"
$ganttTable.Cell($row,3).Range.Text = "3"
$ganttTable.Cell($row,4).Shading.BackgroundPatternColor = -16776961 # Blue

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Permission Analysis"
$ganttTable.Cell($row,3).Range.Text = "2"
$ganttTable.Cell($row,4).Shading.BackgroundPatternColor = -16776961

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Licensing Setup"
$ganttTable.Cell($row,3).Range.Text = "1"
$ganttTable.Cell($row,4).Shading.BackgroundPatternColor = -16776961

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Stakeholder Interviews"
$ganttTable.Cell($row,3).Range.Text = "3"
$ganttTable.Cell($row,4).Shading.BackgroundPatternColor = -16776961
$ganttTable.Cell($row,5).Shading.BackgroundPatternColor = -16776961

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Success Metrics"
$ganttTable.Cell($row,3).Range.Text = "1"
$ganttTable.Cell($row,5).Shading.BackgroundPatternColor = -16776961

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Architecture Design"
$ganttTable.Cell($row,3).Range.Text = "2"
$ganttTable.Cell($row,5).Shading.BackgroundPatternColor = -16776961

# Phase 2: Environment Setup
$row++
$ganttTable.Cell($row,1).Range.Text = "Phase 2"
$ganttTable.Cell($row,2).Range.Text = "Create Environment"
$ganttTable.Cell($row,3).Range.Text = "0.5"
$ganttTable.Cell($row,6).Shading.BackgroundPatternColor = -16744448 # Green

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "SharePoint Connector"
$ganttTable.Cell($row,3).Range.Text = "2"
$ganttTable.Cell($row,6).Shading.BackgroundPatternColor = -16744448

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Document Indexing"
$ganttTable.Cell($row,3).Range.Text = "1.5"
$ganttTable.Cell($row,6).Shading.BackgroundPatternColor = -16744448

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Test Connectivity"
$ganttTable.Cell($row,3).Range.Text = "1"
$ganttTable.Cell($row,6).Shading.BackgroundPatternColor = -16744448

# Phase 3: Development
$row++
$ganttTable.Cell($row,1).Range.Text = "Phase 3"
$ganttTable.Cell($row,2).Range.Text = "Base Gen AI Topic"
$ganttTable.Cell($row,3).Range.Text = "2"
$ganttTable.Cell($row,7).Shading.BackgroundPatternColor = -16737997 # Orange

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Conversation Flows"
$ganttTable.Cell($row,3).Range.Text = "3"
$ganttTable.Cell($row,7).Shading.BackgroundPatternColor = -16737997
$ganttTable.Cell($row,8).Shading.BackgroundPatternColor = -16737997

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Document Citations"
$ganttTable.Cell($row,3).Range.Text = "1.5"
$ganttTable.Cell($row,7).Shading.BackgroundPatternColor = -16737997

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Content Moderation"
$ganttTable.Cell($row,3).Range.Text = "1"
$ganttTable.Cell($row,8).Shading.BackgroundPatternColor = -16737997

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Error Handling"
$ganttTable.Cell($row,3).Range.Text = "1.5"
$ganttTable.Cell($row,8).Shading.BackgroundPatternColor = -16737997

# Phase 4: Testing
$row++
$ganttTable.Cell($row,1).Range.Text = "Phase 4"
$ganttTable.Cell($row,2).Range.Text = "Internal Testing"
$ganttTable.Cell($row,3).Range.Text = "2"
$ganttTable.Cell($row,9).Shading.BackgroundPatternColor = -11489280 # Purple

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Permission Testing"
$ganttTable.Cell($row,3).Range.Text = "2"
$ganttTable.Cell($row,9).Shading.BackgroundPatternColor = -11489280

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Pilot User Testing"
$ganttTable.Cell($row,3).Range.Text = "5"
$ganttTable.Cell($row,9).Shading.BackgroundPatternColor = -11489280
$ganttTable.Cell($row,10).Shading.BackgroundPatternColor = -11489280

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Quality Refinement"
$ganttTable.Cell($row,3).Range.Text = "3"
$ganttTable.Cell($row,10).Shading.BackgroundPatternColor = -11489280

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Performance Testing"
$ganttTable.Cell($row,3).Range.Text = "1"
$ganttTable.Cell($row,10).Shading.BackgroundPatternColor = -11489280

# Phase 5: Deployment
$row++
$ganttTable.Cell($row,1).Range.Text = "Phase 5"
$ganttTable.Cell($row,2).Range.Text = "Production Environment"
$ganttTable.Cell($row,3).Range.Text = "0.5"
$ganttTable.Cell($row,11).Shading.BackgroundPatternColor = -65536 # Red

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Teams/Intranet Integration"
$ganttTable.Cell($row,3).Range.Text = "2"
$ganttTable.Cell($row,11).Shading.BackgroundPatternColor = -65536

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Documentation"
$ganttTable.Cell($row,3).Range.Text = "2"
$ganttTable.Cell($row,11).Shading.BackgroundPatternColor = -65536

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Training Sessions"
$ganttTable.Cell($row,3).Range.Text = "2"
$ganttTable.Cell($row,11).Shading.BackgroundPatternColor = -65536

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Phased Rollout"
$ganttTable.Cell($row,3).Range.Text = "1"
$ganttTable.Cell($row,11).Shading.BackgroundPatternColor = -65536

# Phase 6: Post-Launch
$row++
$ganttTable.Cell($row,1).Range.Text = "Phase 6"
$ganttTable.Cell($row,2).Range.Text = "Monitor Analytics"
$ganttTable.Cell($row,3).Range.Text = "Ongoing"
$ganttTable.Cell($row,12).Shading.BackgroundPatternColor = -16711681 # Cyan
$ganttTable.Cell($row,13).Shading.BackgroundPatternColor = -16711681

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "User Support"
$ganttTable.Cell($row,3).Range.Text = "5"
$ganttTable.Cell($row,12).Shading.BackgroundPatternColor = -16711681
$ganttTable.Cell($row,13).Shading.BackgroundPatternColor = -16711681

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Iterative Improvements"
$ganttTable.Cell($row,3).Range.Text = "3"
$ganttTable.Cell($row,12).Shading.BackgroundPatternColor = -16711681

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Refresh Process Setup"
$ganttTable.Cell($row,3).Range.Text = "1"
$ganttTable.Cell($row,12).Shading.BackgroundPatternColor = -16711681

$row++
$ganttTable.Cell($row,1).Range.Text = ""
$ganttTable.Cell($row,2).Range.Text = "Final Report"
$ganttTable.Cell($row,3).Range.Text = "1"
$ganttTable.Cell($row,13).Shading.BackgroundPatternColor = -16711681

# Adjust column widths
$ganttTable.Columns.Item(1).Width = 60
$ganttTable.Columns.Item(2).Width = 140
$ganttTable.Columns.Item(3).Width = 40
for ($i = 4; $i -le 13; $i++) {
    $ganttTable.Columns.Item($i).Width = 35
}

$selection.EndKey(6) # Move to end of document
$selection.TypeParagraph()
$selection.TypeParagraph()

# Legend
$selection.Font.Size = 11
$selection.Font.Bold = $true
$selection.TypeText("Legend:")
$selection.TypeParagraph()
$selection.Font.Bold = $false
$selection.TypeText("  Blue = Discovery & Planning  |  Green = Environment Setup  |  Orange = Development")
$selection.TypeParagraph()
$selection.TypeText("  Purple = Testing  |  Red = Deployment  |  Cyan = Post-Launch Support")
$selection.TypeParagraph()

# Page break for detailed plan
$selection.InsertBreak(7)

# Detailed Phase Breakdowns
$selection.Font.Size = 18
$selection.Font.Bold = $true
$selection.TypeText("Detailed Project Plan")
$selection.TypeParagraph()
$selection.TypeParagraph()

# Phase 1 Details
$selection.Font.Size = 14
$selection.Font.Bold = $true
$selection.Font.Color = -16776961
$selection.TypeText("Phase 1: Discovery & Planning (Week 1-2)")
$selection.TypeParagraph()
$selection.Font.Color = 0
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()

$phase1Table = $selection.Tables.Add($selection.Range, 7, 5)
$phase1Table.Borders.Enable = $true
$phase1Table.Style = "Grid Table 4"

$phase1Table.Cell(1,1).Range.Text = "Task ID"
$phase1Table.Cell(1,2).Range.Text = "Task Name"
$phase1Table.Cell(1,3).Range.Text = "Duration"
$phase1Table.Cell(1,4).Range.Text = "Dependencies"
$phase1Table.Cell(1,5).Range.Text = "Effort Justification"

$phase1Table.Cell(2,1).Range.Text = "1.1"
$phase1Table.Cell(2,2).Range.Text = "SharePoint Site & Document Audit"
$phase1Table.Cell(2,3).Range.Text = "3 days"
$phase1Table.Cell(2,4).Range.Text = "None"
$phase1Table.Cell(2,5).Range.Text = "• Catalog all 50 documents across sites
• Identify folder structures
• Coordinate with site owners
• Document access patterns"

$phase1Table.Cell(3,1).Range.Text = "1.2"
$phase1Table.Cell(3,2).Range.Text = "Permission & Security Analysis"
$phase1Table.Cell(3,3).Range.Text = "2 days"
$phase1Table.Cell(3,4).Range.Text = "1.1"
$phase1Table.Cell(3,5).Range.Text = "• Map permission groups
• Identify sensitive classifications
• Ensure compliance
• Test permission inheritance"

$phase1Table.Cell(4,1).Range.Text = "1.3"
$phase1Table.Cell(4,2).Range.Text = "Copilot Studio Licensing Setup"
$phase1Table.Cell(4,3).Range.Text = "1 day"
$phase1Table.Cell(4,4).Range.Text = "None"
$phase1Table.Cell(4,5).Range.Text = "• Verify licenses
• Request additional if needed
• Set up dev environment
• Handle procurement"

$phase1Table.Cell(5,1).Range.Text = "1.4"
$phase1Table.Cell(5,2).Range.Text = "Stakeholder Interviews"
$phase1Table.Cell(5,3).Range.Text = "3 days"
$phase1Table.Cell(5,4).Range.Text = "None"
$phase1Table.Cell(5,5).Range.Text = "• Interview 8-10 team members
• Understand common questions
• Document pain points
• Prioritize features"

$phase1Table.Cell(6,1).Range.Text = "1.5"
$phase1Table.Cell(6,2).Range.Text = "Success Metrics Definition"
$phase1Table.Cell(6,3).Range.Text = "1 day"
$phase1Table.Cell(6,4).Range.Text = "1.4"
$phase1Table.Cell(6,5).Range.Text = "• Define success criteria
• Establish baseline metrics
• Set adoption targets
• Create tracking mechanisms"

$phase1Table.Cell(7,1).Range.Text = "1.6"
$phase1Table.Cell(7,2).Range.Text = "Technical Architecture Design"
$phase1Table.Cell(7,3).Range.Text = "2 days"
$phase1Table.Cell(7,4).Range.Text = "1.1, 1.2, 1.3"
$phase1Table.Cell(7,5).Range.Text = "• Design data source connections
• Plan error handling
• Determine integration points
• Document security approach"

$phase1Table.Columns.Item(1).Width = 50
$phase1Table.Columns.Item(2).Width = 140
$phase1Table.Columns.Item(3).Width = 60
$phase1Table.Columns.Item(4).Width = 80
$phase1Table.Columns.Item(5).Width = 180

$selection.EndKey(6)
$selection.TypeParagraph()
$selection.TypeParagraph()

# Phase 2 Details
$selection.Font.Size = 14
$selection.Font.Bold = $true
$selection.Font.Color = -16744448
$selection.TypeText("Phase 2: Environment Setup & Configuration (Week 3)")
$selection.TypeParagraph()
$selection.Font.Color = 0
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()

$phase2Table = $selection.Tables.Add($selection.Range, 5, 5)
$phase2Table.Borders.Enable = $true
$phase2Table.Style = "Grid Table 4"

$phase2Table.Cell(1,1).Range.Text = "Task ID"
$phase2Table.Cell(1,2).Range.Text = "Task Name"
$phase2Table.Cell(1,3).Range.Text = "Duration"
$phase2Table.Cell(1,4).Range.Text = "Dependencies"
$phase2Table.Cell(1,5).Range.Text = "Effort Justification"

$phase2Table.Cell(2,1).Range.Text = "2.1"
$phase2Table.Cell(2,2).Range.Text = "Create Copilot Studio Environment"
$phase2Table.Cell(2,3).Range.Text = "0.5 day"
$phase2Table.Cell(2,4).Range.Text = "1.6"
$phase2Table.Cell(2,5).Range.Text = "• Set up dev environment
• Configure basic settings
• Straightforward process
• Requires careful configuration"

$phase2Table.Cell(3,1).Range.Text = "2.2"
$phase2Table.Cell(3,2).Range.Text = "SharePoint Connector Configuration"
$phase2Table.Cell(3,3).Range.Text = "2 days"
$phase2Table.Cell(3,4).Range.Text = "2.1, 1.1"
$phase2Table.Cell(3,5).Range.Text = "• Connect to multiple sites
• Configure OAuth/service account
• Test connectivity
• Troubleshoot permissions"

$phase2Table.Cell(4,1).Range.Text = "2.3"
$phase2Table.Cell(4,2).Range.Text = "Document Indexing Setup"
$phase2Table.Cell(4,3).Range.Text = "1.5 days"
$phase2Table.Cell(4,4).Range.Text = "2.2"
$phase2Table.Cell(4,5).Range.Text = "• Configure folders to index
• Set up refresh schedules
• Exclude non-relevant files
• Initial indexing takes time"

$phase2Table.Cell(5,1).Range.Text = "2.4"
$phase2Table.Cell(5,2).Range.Text = "Test Data Source Connectivity"
$phase2Table.Cell(5,3).Range.Text = "1 day"
$phase2Table.Cell(5,4).Range.Text = "2.3"
$phase2Table.Cell(5,5).Range.Text = "• Verify document access
• Test different permission levels
• Validate folder hierarchy
• Document issues"

$phase2Table.Columns.Item(1).Width = 50
$phase2Table.Columns.Item(2).Width = 140
$phase2Table.Columns.Item(3).Width = 60
$phase2Table.Columns.Item(4).Width = 80
$phase2Table.Columns.Item(5).Width = 180

$selection.EndKey(6)
$selection.TypeParagraph()
$selection.TypeParagraph()

# Phase 3 Details
$selection.Font.Size = 14
$selection.Font.Bold = $true
$selection.Font.Color = -16737997
$selection.TypeText("Phase 3: Copilot Development (Week 4-5)")
$selection.TypeParagraph()
$selection.Font.Color = 0
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()

$phase3Table = $selection.Tables.Add($selection.Range, 6, 5)
$phase3Table.Borders.Enable = $true
$phase3Table.Style = "Grid Table 4"

$phase3Table.Cell(1,1).Range.Text = "Task ID"
$phase3Table.Cell(1,2).Range.Text = "Task Name"
$phase3Table.Cell(1,3).Range.Text = "Duration"
$phase3Table.Cell(1,4).Range.Text = "Dependencies"
$phase3Table.Cell(1,5).Range.Text = "Effort Justification"

$phase3Table.Cell(2,1).Range.Text = "3.1"
$phase3Table.Cell(2,2).Range.Text = "Create Base Generative Answers"
$phase3Table.Cell(2,3).Range.Text = "2 days"
$phase3Table.Cell(2,4).Range.Text = "2.4"
$phase3Table.Cell(2,5).Range.Text = "• Configure AI model settings
• Prompt engineering for BCP
• Define response format
• Test with sample queries"

$phase3Table.Cell(3,1).Range.Text = "3.2"
$phase3Table.Cell(3,2).Range.Text = "Design Conversation Flow"
$phase3Table.Cell(3,3).Range.Text = "3 days"
$phase3Table.Cell(3,4).Range.Text = "3.1, 1.4"
$phase3Table.Cell(3,5).Range.Text = "• Create guided topics
• Design conversation trees
• Build fallback mechanisms
• Multiple iterations needed"

$phase3Table.Cell(4,1).Range.Text = "3.3"
$phase3Table.Cell(4,2).Range.Text = "Implement Document Citation"
$phase3Table.Cell(4,3).Range.Text = "1.5 days"
$phase3Table.Cell(4,4).Range.Text = "3.1"
$phase3Table.Cell(4,5).Range.Text = "• Configure citation format
• Include SharePoint links
• Test user access to sources
• Format for readability"

$phase3Table.Cell(5,1).Range.Text = "3.4"
$phase3Table.Cell(5,2).Range.Text = "Content Moderation & Filters"
$phase3Table.Cell(5,3).Range.Text = "1 day"
$phase3Table.Cell(5,4).Range.Text = "3.1"
$phase3Table.Cell(5,5).Range.Text = "• Set up moderation policies
• Configure PII detection
• Test with edge cases
• Ensure compliance"

$phase3Table.Cell(6,1).Range.Text = "3.5"
$phase3Table.Cell(6,2).Range.Text = "Error Handling & Fallback"
$phase3Table.Cell(6,3).Range.Text = "1.5 days"
$phase3Table.Cell(6,4).Range.Text = "3.2"
$phase3Table.Cell(6,5).Range.Text = "• Design unanswerable responses
• Create escalation paths
• Handle connectivity issues
• Test failure scenarios"

$phase3Table.Columns.Item(1).Width = 50
$phase3Table.Columns.Item(2).Width = 140
$phase3Table.Columns.Item(3).Width = 60
$phase3Table.Columns.Item(4).Width = 80
$phase3Table.Columns.Item(5).Width = 180

$selection.EndKey(6)
$selection.InsertBreak(7)

# Phase 4 Details
$selection.Font.Size = 14
$selection.Font.Bold = $true
$selection.Font.Color = -11489280
$selection.TypeText("Phase 4: Testing & Refinement (Week 6-7)")
$selection.TypeParagraph()
$selection.Font.Color = 0
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()

$phase4Table = $selection.Tables.Add($selection.Range, 6, 5)
$phase4Table.Borders.Enable = $true
$phase4Table.Style = "Grid Table 4"

$phase4Table.Cell(1,1).Range.Text = "Task ID"
$phase4Table.Cell(1,2).Range.Text = "Task Name"
$phase4Table.Cell(1,3).Range.Text = "Duration"
$phase4Table.Cell(1,4).Range.Text = "Dependencies"
$phase4Table.Cell(1,5).Range.Text = "Effort Justification"

$phase4Table.Cell(2,1).Range.Text = "4.1"
$phase4Table.Cell(2,2).Range.Text = "Internal Testing"
$phase4Table.Cell(2,3).Range.Text = "2 days"
$phase4Table.Cell(2,4).Range.Text = "3.5"
$phase4Table.Cell(2,5).Range.Text = "• Test 30-50 BCP questions
• Verify response accuracy
• Check citation correctness
• Document issues"

$phase4Table.Cell(3,1).Range.Text = "4.2"
$phase4Table.Cell(3,2).Range.Text = "Permission Testing"
$phase4Table.Cell(3,3).Range.Text = "2 days"
$phase4Table.Cell(3,4).Range.Text = "4.1"
$phase4Table.Cell(3,5).Range.Text = "• Test different access levels
• Verify authorized content only
• Test edge cases
• Multiple test accounts"

$phase4Table.Cell(4,1).Range.Text = "4.3"
$phase4Table.Cell(4,2).Range.Text = "Pilot User Group Testing"
$phase4Table.Cell(4,3).Range.Text = "5 days"
$phase4Table.Cell(4,4).Range.Text = "4.2"
$phase4Table.Cell(4,5).Range.Text = "• Select diverse pilot group
• Provide training
• Collect feedback over 3-5 days
• Real scenario testing needed"

$phase4Table.Cell(5,1).Range.Text = "4.4"
$phase4Table.Cell(5,2).Range.Text = "Response Quality Refinement"
$phase4Table.Cell(5,3).Range.Text = "3 days"
$phase4Table.Cell(5,4).Range.Text = "4.3"
$phase4Table.Cell(5,5).Range.Text = "• Analyze pilot feedback
• Adjust prompt engineering
• Improve topic flows
• Iterative refinement"

$phase4Table.Cell(6,1).Range.Text = "4.5"
$phase4Table.Cell(6,2).Range.Text = "Performance & Load Testing"
$phase4Table.Cell(6,3).Range.Text = "1 day"
$phase4Table.Cell(6,4).Range.Text = "4.4"
$phase4Table.Cell(6,5).Range.Text = "• Test concurrent users
• Verify response times
• Check indexing refresh
• Identify bottlenecks"

$phase4Table.Columns.Item(1).Width = 50
$phase4Table.Columns.Item(2).Width = 140
$phase4Table.Columns.Item(3).Width = 60
$phase4Table.Columns.Item(4).Width = 80
$phase4Table.Columns.Item(5).Width = 180

$selection.EndKey(6)
$selection.TypeParagraph()
$selection.TypeParagraph()

# Phase 5 Details
$selection.Font.Size = 14
$selection.Font.Bold = $true
$selection.Font.Color = -65536
$selection.TypeText("Phase 5: Deployment & Rollout (Week 8)")
$selection.TypeParagraph()
$selection.Font.Color = 0
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()

$phase5Table = $selection.Tables.Add($selection.Range, 6, 5)
$phase5Table.Borders.Enable = $true
$phase5Table.Style = "Grid Table 4"

$phase5Table.Cell(1,1).Range.Text = "Task ID"
$phase5Table.Cell(1,2).Range.Text = "Task Name"
$phase5Table.Cell(1,3).Range.Text = "Duration"
$phase5Table.Cell(1,4).Range.Text = "Dependencies"
$phase5Table.Cell(1,5).Range.Text = "Effort Justification"

$phase5Table.Cell(2,1).Range.Text = "5.1"
$phase5Table.Cell(2,2).Range.Text = "Create Production Environment"
$phase5Table.Cell(2,3).Range.Text = "0.5 day"
$phase5Table.Cell(2,4).Range.Text = "4.5"
$phase5Table.Cell(2,5).Range.Text = "• Clone dev to prod
• Verify all connections
• Quick but careful validation
• Production settings"

$phase5Table.Cell(3,1).Range.Text = "5.2"
$phase5Table.Cell(3,2).Range.Text = "Integration with Teams/Intranet"
$phase5Table.Cell(3,3).Range.Text = "2 days"
$phase5Table.Cell(3,4).Range.Text = "5.1"
$phase5Table.Cell(3,5).Range.Text = "• Deploy as Teams app
• Add to intranet if needed
• Configure SSO
• Test integrations"

$phase5Table.Cell(4,1).Range.Text = "5.3"
$phase5Table.Cell(4,2).Range.Text = "User Documentation"
$phase5Table.Cell(4,3).Range.Text = "2 days"
$phase5Table.Cell(4,4).Range.Text = "4.5"
$phase5Table.Cell(4,5).Range.Text = "• Create user guide
• Develop quick reference
• Record demo video
• Prepare FAQ"

$phase5Table.Cell(5,1).Range.Text = "5.4"
$phase5Table.Cell(5,2).Range.Text = "Conduct Training Sessions"
$phase5Table.Cell(5,3).Range.Text = "2 days"
$phase5Table.Cell(5,4).Range.Text = "5.3"
$phase5Table.Cell(5,5).Range.Text = "• Schedule 2-3 sessions
• Live demos and Q&A
• Hands-on practice
• Record for reference"

$phase5Table.Cell(6,1).Range.Text = "5.5"
$phase5Table.Cell(6,2).Range.Text = "Phased Rollout to Full Team"
$phase5Table.Cell(6,3).Range.Text = "1 day"
$phase5Table.Cell(6,4).Range.Text = "5.2, 5.4"
$phase5Table.Cell(6,5).Range.Text = "• Grant access in waves
• Monitor for issues
• Provide initial support
• Collect early feedback"

$phase5Table.Columns.Item(1).Width = 50
$phase5Table.Columns.Item(2).Width = 140
$phase5Table.Columns.Item(3).Width = 60
$phase5Table.Columns.Item(4).Width = 80
$phase5Table.Columns.Item(5).Width = 180

$selection.EndKey(6)
$selection.TypeParagraph()
$selection.TypeParagraph()

# Phase 6 Details
$selection.Font.Size = 14
$selection.Font.Bold = $true
$selection.Font.Color = -16711681
$selection.TypeText("Phase 6: Post-Launch Support & Optimization (Week 9-10)")
$selection.TypeParagraph()
$selection.Font.Color = 0
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()

$phase6Table = $selection.Tables.Add($selection.Range, 6, 5)
$phase6Table.Borders.Enable = $true
$phase6Table.Style = "Grid Table 4"

$phase6Table.Cell(1,1).Range.Text = "Task ID"
$phase6Table.Cell(1,2).Range.Text = "Task Name"
$phase6Table.Cell(1,3).Range.Text = "Duration"
$phase6Table.Cell(1,4).Range.Text = "Dependencies"
$phase6Table.Cell(1,5).Range.Text = "Effort Justification"

$phase6Table.Cell(2,1).Range.Text = "6.1"
$phase6Table.Cell(2,2).Range.Text = "Monitor Usage & Analytics"
$phase6Table.Cell(2,3).Range.Text = "Ongoing"
$phase6Table.Cell(2,4).Range.Text = "5.5"
$phase6Table.Cell(2,5).Range.Text = "• Daily monitoring 2 weeks
• Track adoption metrics
• Identify problem queries
• Continuous activity"

$phase6Table.Cell(3,1).Range.Text = "6.2"
$phase6Table.Cell(3,2).Range.Text = "User Support & Issue Resolution"
$phase6Table.Cell(3,3).Range.Text = "5 days"
$phase6Table.Cell(3,4).Range.Text = "5.5"
$phase6Table.Cell(3,5).Range.Text = "• Respond to user questions
• Fix bugs and issues
• Provide helpdesk support
• Higher initial load"

$phase6Table.Cell(4,1).Range.Text = "6.3"
$phase6Table.Cell(4,2).Range.Text = "Iterative Improvements"
$phase6Table.Cell(4,3).Range.Text = "3 days"
$phase6Table.Cell(4,4).Range.Text = "6.1"
$phase6Table.Cell(4,5).Range.Text = "• Analyze common queries
• Add new topics
• Improve low-confidence responses
• Optimize flows"

$phase6Table.Cell(5,1).Range.Text = "6.4"
$phase6Table.Cell(5,2).Range.Text = "Document Refresh Process"
$phase6Table.Cell(5,3).Range.Text = "1 day"
$phase6Table.Cell(5,4).Range.Text = "6.1"
$phase6Table.Cell(5,5).Range.Text = "• Process for new documents
• Configure change alerts
• Test refresh/reindexing
• Create maintenance runbook"

$phase6Table.Cell(6,1).Range.Text = "6.5"
$phase6Table.Cell(6,2).Range.Text = "Final Report & Handoff"
$phase6Table.Cell(6,3).Range.Text = "1 day"
$phase6Table.Cell(6,4).Range.Text = "6.3, 6.4"
$phase6Table.Cell(6,5).Range.Text = "• Compare to success metrics
• Document lessons learned
• Create maintenance guide
• Handoff to operations"

$phase6Table.Columns.Item(1).Width = 50
$phase6Table.Columns.Item(2).Width = 140
$phase6Table.Columns.Item(3).Width = 60
$phase6Table.Columns.Item(4).Width = 80
$phase6Table.Columns.Item(5).Width = 180

$selection.EndKey(6)
$selection.InsertBreak(7)

# Risk Management
$selection.Font.Size = 18
$selection.Font.Bold = $true
$selection.TypeText("Risk Management & Mitigation")
$selection.TypeParagraph()
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()

$riskTable = $selection.Tables.Add($selection.Range, 6, 4)
$riskTable.Borders.Enable = $true
$riskTable.Style = "Grid Table 4"

$riskTable.Cell(1,1).Range.Text = "Risk"
$riskTable.Cell(1,2).Range.Text = "Impact"
$riskTable.Cell(1,3).Range.Text = "Probability"
$riskTable.Cell(1,4).Range.Text = "Mitigation Strategy"

$riskTable.Cell(2,1).Range.Text = "SharePoint Permission Complexity"
$riskTable.Cell(2,2).Range.Text = "High"
$riskTable.Cell(2,3).Range.Text = "Medium"
$riskTable.Cell(2,4).Range.Text = "• Add 2-3 days buffer to Phase 1-2
• Early permission testing
• Engage SharePoint admin early
• Document permission structure thoroughly"

$riskTable.Cell(3,1).Range.Text = "Licensing Procurement Delays"
$riskTable.Cell(3,2).Range.Text = "High"
$riskTable.Cell(3,3).Range.Text = "Medium"
$riskTable.Cell(3,4).Range.Text = "• Start licensing request in Week 1
• Identify procurement lead time
• Have backup approval path
• Can add 1-2 weeks to timeline"

$riskTable.Cell(4,1).Range.Text = "Pilot Feedback Requires Major Changes"
$riskTable.Cell(4,2).Range.Text = "Medium"
$riskTable.Cell(4,3).Range.Text = "Medium"
$riskTable.Cell(4,4).Range.Text = "• Select representative pilot group
• Set clear expectations
• May extend Week 6-7 by 3-5 days
• Prioritize critical feedback"

$riskTable.Cell(5,1).Range.Text = "Document Quality Issues"
$riskTable.Cell(5,2).Range.Text = "Medium"
$riskTable.Cell(5,3).Range.Text = "Low"
$riskTable.Cell(5,4).Range.Text = "• Identify during Phase 1 audit
• May need document cleanup (add 1 week)
• Work with content owners
• Use OCR if needed for scans"

$riskTable.Cell(6,1).Range.Text = "User Adoption Challenges"
$riskTable.Cell(6,2).Range.Text = "Medium"
$riskTable.Cell(6,3).Range.Text = "Low"
$riskTable.Cell(6,4).Range.Text = "• Comprehensive training program
• Champions/advocates approach
• Collect and act on feedback quickly
• Show quick wins early"

$riskTable.Columns.Item(1).Width = 130
$riskTable.Columns.Item(2).Width = 60
$riskTable.Columns.Item(3).Width = 70
$riskTable.Columns.Item(4).Width = 250

$selection.EndKey(6)
$selection.TypeParagraph()
$selection.TypeParagraph()

# Success Criteria
$selection.Font.Size = 18
$selection.Font.Bold = $true
$selection.TypeText("Success Criteria & KPIs")
$selection.TypeParagraph()
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()

$kpiTable = $selection.Tables.Add($selection.Range, 6, 3)
$kpiTable.Borders.Enable = $true
$kpiTable.Style = "Grid Table 4"

$kpiTable.Cell(1,1).Range.Text = "KPI Category"
$kpiTable.Cell(1,2).Range.Text = "Metric"
$kpiTable.Cell(1,3).Range.Text = "Target"

$kpiTable.Cell(2,1).Range.Text = "User Adoption"
$kpiTable.Cell(2,2).Range.Text = "Active users in first month"
$kpiTable.Cell(2,3).Range.Text = "80% of BCP team"

$kpiTable.Cell(3,1).Range.Text = "Response Quality"
$kpiTable.Cell(3,2).Range.Text = "User satisfaction rating"
$kpiTable.Cell(3,3).Range.Text = "4.0+ out of 5.0"

$kpiTable.Cell(4,1).Range.Text = "Efficiency Gain"
$kpiTable.Cell(4,2).Range.Text = "Time saved per query vs manual search"
$kpiTable.Cell(4,3).Range.Text = "70% reduction (5 min → 1.5 min)"

$kpiTable.Cell(5,1).Range.Text = "Accuracy"
$kpiTable.Cell(5,2).Range.Text = "Correct answers with proper citations"
$kpiTable.Cell(5,3).Range.Text = "85%+ accuracy rate"

$kpiTable.Cell(6,1).Range.Text = "Performance"
$kpiTable.Cell(6,2).Range.Text = "Average response time"
$kpiTable.Cell(6,3).Range.Text = "< 5 seconds per query"

$kpiTable.Columns.Item(1).Width = 120
$kpiTable.Columns.Item(2).Width = 200
$kpiTable.Columns.Item(3).Width = 190

$selection.EndKey(6)
$selection.TypeParagraph()
$selection.TypeParagraph()

# Next Steps
$selection.Font.Size = 18
$selection.Font.Bold = $true
$selection.TypeText("Immediate Next Steps")
$selection.TypeParagraph()
$selection.Font.Size = 11
$selection.Font.Bold = $false
$selection.TypeParagraph()

$selection.TypeText("1. Secure Executive Approval")
$selection.TypeParagraph()
$selection.TypeText("   • Present project plan to leadership")
$selection.TypeParagraph()
$selection.TypeText("   • Obtain budget approval")
$selection.TypeParagraph()
$selection.TypeText("   • Confirm resource allocation")
$selection.TypeParagraph()
$selection.TypeParagraph()

$selection.TypeText("2. Initiate Licensing Process (CRITICAL)")
$selection.TypeParagraph()
$selection.TypeText("   • Verify current Copilot Studio licenses")
$selection.TypeParagraph()
$selection.TypeText("   • Submit procurement request for additional licenses")
$selection.TypeParagraph()
$selection.TypeText("   • Identify lead time for license activation")
$selection.TypeParagraph()
$selection.TypeParagraph()

$selection.TypeText("3. Assemble Project Team")
$selection.TypeParagraph()
$selection.TypeText("   • Assign Project Manager")
$selection.TypeParagraph()
$selection.TypeText("   • Identify Copilot Studio Developer")
$selection.TypeParagraph()
$selection.TypeText("   • Engage SharePoint Administrator")
$selection.TypeParagraph()
$selection.TypeText("   • Designate BCP Subject Matter Expert")
$selection.TypeParagraph()
$selection.TypeParagraph()

$selection.TypeText("4. Schedule Kickoff Meeting")
$selection.TypeParagraph()
$selection.TypeText("   • Set project start date")
$selection.TypeParagraph()
$selection.TypeText("   • Align team calendars")
$selection.TypeParagraph()
$selection.TypeText("   • Review project plan with team")
$selection.TypeParagraph()
$selection.TypeParagraph()

$selection.TypeText("5. Begin Phase 1 Activities")
$selection.TypeParagraph()
$selection.TypeText("   • Start SharePoint site audit")
$selection.TypeParagraph()
$selection.TypeText("   • Schedule stakeholder interviews")
$selection.TypeParagraph()
$selection.TypeText("   • Initiate licensing setup")
$selection.TypeParagraph()

# Save document
$filePath = "c:\Users\aniru\source\repos\BusinessContinuity\ProjectPlan\BCP_Copilot_Project_Plan.docx"
$document.SaveAs([ref]$filePath)
$document.Close()
$doc.Quit()

Write-Host "Document created successfully at: $filePath" -ForegroundColor Green
