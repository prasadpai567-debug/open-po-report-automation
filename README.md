# Open PO Report Automation

> [!IMPORTANT]
> Repository status (2026-11-01): this public repository currently contains documentation only: .gitignore, README.md, and ANALYSIS_REPORT.md. The Excel workbook and VBA source described below are not committed, so the automation cannot be run from this repository. The implementation details are based on a separately inspected workbook and are not verifiable from the current GitHub tree.

## What is verifiable here

The GitHub tree verifies the documentation files and publication safeguards. It does not verify SAP, Excel, VBA, or Outlook behavior. Do not treat this repository as a runnable release until a sanitized workbook or source package has been added and reviewed.

## Project overview

This repository documents an Excel VBA workflow intended to retrieve an open purchase-order report from SAP, export the result as an Excel file, and prepare an Outlook email draft for human review.

The documented workflow is designed to reduce repeated procurement reporting steps while keeping final email delivery under user control.

## Business problem solved

Without automation, a user must repeatedly navigate SAP, apply the same report variant, enter dates, export the report, locate or rename the output file, and create a standard email. The documented workflow turns that sequence into a repeatable button-driven process.

No quantified time savings or error-rate improvements are claimed because the repository contains no measured operational results.

## Documented workflow capabilities

The following behavior is recorded from a separately inspected workbook. It is not verifiable from the current GitHub tree until a sanitized workbook or source package is added.

- Enter report start and end dates in worksheet cells B2 and B3.
- Connect to an already logged-in SAP GUI session.
- Open SAP transaction ME2N and apply a configured saved variant.
- Enter a document-date range and export the report as an Excel file.
- Save the output under a local temporary folder using a date-based filename.
- Create an Outlook draft with the exported report attached.
- Display the draft for review; the documented behavior does not send the email automatically.
- Exit with a message when no active SAP session is found.

## Application workflow

1. Enter the start date and end date in the workbook.
2. Click the Run Macro button.
3. Connect to the active SAP GUI session.
4. Open ME2N and apply the saved variant.
5. Enter the date range and execute the report.
6. Export the report to a local file.
7. Create and display an Outlook draft with the report attached.
8. Review the date range, report content, recipients, subject, and attachment before sending manually.

## Inputs and outputs

| Item | Documented behavior |
| --- | --- |
| Start date | Worksheet cell B2 |
| End date | Worksheet cell B3 |
| SAP transaction | ME2N |
| Output format | Excel .xlsx |
| Email client | Microsoft Outlook desktop |
| Email action | Draft displayed for review; not sent automatically |

## Technology stack

- Microsoft Excel macro-enabled workbook (.xlsm)
- Visual Basic for Applications (VBA)
- SAP GUI Scripting through late-bound COM objects
- Microsoft Outlook desktop automation
- Windows local file system

## Intended runtime requirements

- Windows with desktop Microsoft Excel and VBA support
- SAP GUI for Windows with GUI Scripting enabled
- An active SAP logon session with access to ME2N and the configured variant
- Microsoft Outlook desktop with a usable profile
- Write access to the approved output folder
- Organization-approved macro execution permissions

## Installation and setup

> These steps describe the intended runtime setup. They are not currently executable from this repository because the workbook is not included.

1. Download or clone the repository.
2. Obtain an approved sanitized workbook from the project owner.
3. Open the workbook in desktop Excel and follow the organization's macro-security policy.
4. Confirm that the expected date inputs exist in B2 and B3.
5. Confirm that the SAP transaction and saved variant are valid for the target environment.
6. Confirm that the approved output folder is writable.
7. Log in to SAP before running the macro.
8. Open Outlook if the email-draft step is required.

## Usage

1. Enter the report start date in B2.
2. Enter the report end date in B3.
3. Leave the required SAP GUI session active.
4. Click Run Macro.
5. Wait for the SAP export to complete.
6. Review the Outlook draft and attachment.
7. Verify all report and recipient details before sending manually.

## Validation and testing status

- Repository tree and metadata review: completed.
- Static workbook/VBA review: recorded in ANALYSIS_REPORT.md, but the inspected workbook is not committed here.
- Automated tests: none present.
- Excel, SAP GUI, and Outlook runtime testing: not performed.
- No production data was modified.

## Security and publication checklist

Before adding any workbook or source file:

- Use a sanitized copy only.
- Remove recipient, CC, sender-on-behalf, company, plant, and internal mailbox values.
- Remove internal SharePoint, OneDrive, and local storage paths.
- Remove personal and organization metadata from the workbook properties.
- Remove production dates and any supplier, purchase-order, material, price, or report data.
- Do not commit generated Excel reports or email exports.
- Test with synthetic dates and non-production recipients.
- Use an approved signing and deployment process for internal macro distribution.

## Known limitations and operational risks

The separately inspected workbook was reported to have the following limitations:

- SAP GUI control paths are coupled to the current screen layout and configuration.
- The saved variant and output settings are hard-coded.
- Date values are not fully validated before execution.
- Broad error suppression can hide SAP or Outlook failures.
- Export completion and output-file existence are not explicitly verified.
- The first SAP connection and session may not be the intended ones when multiple sessions are open.
- No retry, timeout, execution log, or completion status is currently documented.
- Email content and recipients are not configurable from the worksheet.

## Recommended publication structure

~~~text
open-po-report-automation/
├── README.md
├── ANALYSIS_REPORT.md
├── .gitignore
├── LICENSE                    # only after distribution rights are confirmed
├── workbook/
│   ├── OpenPO_Report_Automation_sanitized.xlsm
│   └── src/Module1.bas
├── docs/
│   ├── architecture.md
│   ├── setup-notes.md
│   └── screenshots/
└── examples/
    └── sample-input-notes.md
~~~

The original production workbook should remain outside the public repository. If it cannot be sanitized safely, publish documentation and synthetic examples instead.

## Screenshots and demonstrations

Use redacted or synthetic examples only. Suggested placeholders:

- Excel input sheet showing B2, B3, and the Run Macro button
- VBA Project Explorer with internal addresses and paths removed
- SAP selection screen with the saved variant redacted
- SAP export dialog using a non-production path
- Outlook draft with recipients redacted

## Future improvements

- Add date validation and clear user-facing error messages.
- Replace broad error suppression with structured error handling and an execution log.
- Verify SAP export completion and output-file existence before creating the draft.
- Use a configurable or user-specific output folder.
- Select the intended SAP connection and session explicitly.
- Add an offline test mode using synthetic or mocked data.
- Digitally sign the VBA project for controlled internal deployment.

## Portfolio summary

Developed a desktop Excel VBA automation workflow for procurement reporting. The documented solution reads a selected date range, drives SAP GUI to retrieve an open purchase-order report, saves the result as an Excel file, and prepares a reviewable Outlook draft with the report attached.

Skills demonstrated: Excel VBA, SAP GUI Scripting, Microsoft Office automation, COM integration, procurement operations, process mapping, documentation, and data confidentiality awareness.

## License and data note

Choose a license only after confirming that the code and workbook may be distributed publicly. The original workbook was reported to contain organization-specific metadata and hard-coded communication details, so it should remain an internal asset until sanitized and approved for release.

For the detailed evidence and limitations, see ANALYSIS_REPORT.md.
