# Training Unit Manager for Process Manager

A PowerShell script for importing and exporting training units to/from Process Manager (Promapp).

## Features

- **Export** training units from Process Manager to CSV
- **Import** training units from CSV to Process Manager
- Automatic pagination handling for large datasets
- Error handling with detailed reporting
- Secure credential input

## Requirements

- PowerShell 5.1 or higher
- Process Manager site access
- Service account with appropriate permissions
- SCIM API key for user lookups

## Usage

### Running the Script

```powershell
.\TrainingUnitManager.ps1
```

### Input Requirements

When you run the script, you'll be prompted for:

1. **Process Manager Site URL** - Format: `https://us.promapp.com/sitename`
2. **Service Account Username** - The username for API authentication
3. **Service Account Password** - Secure password input (not displayed)
4. **SCIM API Key** - API key for SCIM user lookups (not displayed)
5. **Action Selection** - Choose between Export (1) or Import (2)

### Export Functionality

The export feature will:
- Fetch all training units from your Process Manager site
- Retrieve full details for each training unit including:
  - Basic information (Title, Description, Type, etc.)
  - Linked processes with titles and IDs
  - Linked documents
  - Assigned trainees (with username lookup via SCIM API)
- Save data to a CSV file named `TrainingUnits_Export_YYYYMMDD.csv`

**Note:** The export process retrieves trainee UserIds from the Training/Trainee endpoint, then queries the SCIM API to lookup usernames (email addresses) for each trainee. This provides a reliable identifier for future trainee assignment during import.

**Export CSV Columns:**
- Title
- Description
- Type (label value: Course, Online Resource, Document, Face to Face)
- Assessment Label (label value: None, Self Sign Off, Supervisor Sign Off)
- Renew Cycle (integer value)
- Provider
- Linked Processes: Title (semicolon-delimited)
- Linked Processes: uniqueId (semicolon-delimited)
- Linked Documents: Titles (semicolon-delimited)
- Trainees: Usernames (semicolon-delimited, retrieved from SCIM API)

### Import Functionality

The import feature will:
- Read a CSV file with training unit data
- Look up process details for each linked process uniqueId
- Create new training units via the API
- Look up user IDs from usernames via SCIM API
- Assign trainees to the newly created training units
- Skip failed rows and continue processing
- Display a summary report at the end

**Import CSV Columns Required:**
- Title
- Description
- Type (label or integer: see Type Values below)
- Assessment Label (label or integer: see Assessment Method Values below)
- Renew Cycle (integer: see Renew Cycle Values below)
- Provider
- Linked Processes: Title (semicolon-delimited, optional - for reference only)
- Linked Processes: uniqueId (semicolon-delimited, no spaces)
- Linked Documents: Titles (semicolon-delimited, no spaces)
- Trainees: Usernames (semicolon-delimited, optional - usernames/email addresses from SCIM)

## Field Value Reference

### Type Values
You can use either the label or the integer value in your import CSV:

| Label | Integer Value |
|-------|---------------|
| Course | 1 |
| Online Resource | 2 |
| Document | 3 |
| Face to Face | 6 |

**Note:** Export files will use the label format (e.g., "Course"). Import accepts both formats for backward compatibility.

### Assessment Method Values
You can use either the label or the integer value in your import CSV:

| Label | Integer Value |
|-------|---------------|
| None | 0 |
| Self Sign Off | 1 |
| Supervisor Sign Off | 2 |

**Note:** Export files will use the label format (e.g., "Self Sign Off"). Import accepts both formats for backward compatibility.

### Renew Cycle Values
Contact your Process Manager administrator for the specific integer values used in your system. Common examples:
- `1` - Once Only
- `12` - Annually
- `24` - Every 2 Years

## Example Import CSV

See [ImportTemplate.csv](ImportTemplate.csv) for an example CSV file format.

```csv
Title,Description,Type,Assessment Label,Renew Cycle,Provider,Linked Processes: Title,Linked Processes: uniqueId,Linked Documents: Titles,Trainees: Usernames
Safety Training 101,Basic safety training for all employees,Course,Self Sign Off,1,Safety Corp,,3be24da1-4e95-4edb-b94c-f39e14c61081,Safety Manual.pdf;Emergency Procedures.docx,john.doe@example.com;jane.smith@example.com
Advanced Excel Course,Advanced Excel training for data analysts,Online Resource,Supervisor Sign Off,12,Tech Training LLC,,,Excel Guide.xlsx,
```

## Error Handling

### Export Errors
- If authentication fails, the script will exit with an error message
- If a training unit cannot be retrieved, it will be skipped with a warning
- All successfully retrieved units will still be exported

### Import Errors
- The script uses a "skip and continue" approach for failed rows
- Each error is logged with the row number and error message
- A summary report is displayed at the end showing:
  - Total rows processed
  - Number of successful creations
  - Number of failed rows with details

## API Endpoints Used

The script interacts with the following Process Manager API endpoints:

**Authentication:**
- `POST /{tenant}/oauth2/token` - OAuth2 authentication

**Training Unit Operations:**
- `GET /{tenant}/Training/Register/ListPage` - List all training units (paginated)
- `GET /{tenant}/Training/Unit/GetTrainingUnitDetails` - Get full unit details
- `POST /{tenant}/Training/Unit/EditTrainingUnit` - Create/update training unit
- `GET /{tenant}/Training/Trainee` - Get trainees assigned to a unit (paginated)
- `POST /{tenant}/Training/Schedule/SaveSchedule` - Assign trainees to a unit

**Process Operations:**
- `GET /{tenant}/Api/v1/Processes/{processUniqueId}` - Get process details by unique ID

**SCIM API Operations:**
- `GET https://api.promapp.com/api/scim/users/{id}` - Get user by ID
- `GET https://api.promapp.com/api/scim/users?filter={filter}` - Search users by username

## Known Limitations

1. **Document Lookup**: Linked documents are stored by title only. If you need to link documents during import, you'll need to ensure the document titles match exactly.
2. **Trainee Scheduling**: While trainees are assigned during import, advanced scheduling options (custom due dates, supervisor assignment) are not yet supported through the CSV import.

## Troubleshooting

### Authentication Fails
- Verify your site URL is in the correct format: `https://domain.com/tenant`
- Ensure your service account credentials are correct
- Check that the service account has API access enabled

### Process Not Found During Import
- Verify the process uniqueId exists in your system
- Ensure you're using the correct format (GUID with hyphens)
- The process must be published and accessible

### Import Creates Incomplete Records
- Check that all required fields are populated in your CSV
- Verify integer values are correct for Type, Assessment Method, and Renew Cycle
- Ensure semicolon-delimited fields have no spaces after semicolons

### Trainee Assignment Fails
- Verify the usernames in the CSV match exactly with SCIM usernames
- Check that the SCIM API key has permissions to query users
- Ensure users exist in the system before attempting to assign them
- Review the console output for specific user lookup failures

## Security Notes

- Credentials are input securely using PowerShell's `Read-Host -AsSecureString`
- Credentials are only stored in memory during script execution
- The bearer token expires after the duration specified (60000 seconds)
- No credentials are written to disk or logs

## Support

For issues or questions:
1. Check the error messages in the script output
2. Verify your CSV format matches the template
3. Consult your Process Manager administrator for system-specific values
4. Review the API documentation for your Process Manager instance

## Version History

**v2.0** (Current)
- **Trainee Assignment During Import**
  - Import now supports assigning trainees to training units
  - Added SCIM username to UserId lookup function
  - Automatically assigns trainees listed in the CSV after creating training unit
  - Uses SaveSchedule endpoint for trainee assignment
  - Provides detailed feedback on trainee lookup and assignment status
  - Updated CSV format to include "Trainees: Usernames" column

**v1.3**
- **Enhanced Trainee Export with SCIM Integration**
  - Export now retrieves usernames from SCIM API instead of full names
  - Trainees are now exported with their username (email address) for easier import
  - Added SCIM user lookup function using UserId from Training/Trainee endpoint
  - SCIM API is queried by UserId to retrieve userName field
  - Usernames are more suitable for future trainee assignment during import

**v1.2**
- **Enhanced Assessment Label Field Handling**
  - Export now outputs Assessment Label as labels (None, Self Sign Off, Supervisor Sign Off)
  - Import accepts both labels and integer values for backward compatibility
  - Added Assessment conversion helper functions

**v1.1**
- **Enhanced Type Field Handling**
  - Export now outputs Type as labels (Course, Online Resource, Document, Face to Face)
  - Import accepts both labels and integer values for backward compatibility
  - Added Type conversion helper functions

**v1.0**
- Initial release
- Export training units to CSV
- Import training units from CSV
- Process lookup and linking
- Error handling and reporting
