# Feedback Submission SPFx Web Part

## Summary

The **Feedback Submission Solution** is a SharePoint Framework (SPFx) client-side web part designed to dynamically collect user feedback and save it into a SharePoint list. 

The web part pulls up to 5 survey questions from a configurable source list (default: `Feedbackquestions`) and submits the user's responses to a target list (default: `Download_Tracking`). It can also parse a Base64-encrypted `TrackingID` from the page's URL query string, linking the feedback response to a specific transaction or user action.

## Used SharePoint Framework Version

![version](https://img.shields.io/badge/version-1.23.2-green.svg)

## Applies to

- [SharePoint Framework](https://aka.ms/spfx)
- [Microsoft 365 tenant](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)

## Prerequisites

To use this web part, ensure the following SharePoint lists exist in your site:
1. **Target List** (e.g., `Download_Tracking`): Stores the submitted feedback. Must contain fields for the questions, tracking ID, and user acknowledgment.
2. **Questions List** (e.g., `Feedbackquestions`): Stores the questions to be displayed in the web part.

## Solution

| Solution    | Author(s)                                               |
| ----------- | ------------------------------------------------------- |
| FeedbackSubmissionWebPart | Vishnu |

## Version history

| Version | Date             | Comments        |
| ------- | ---------------- | --------------- |
| 1.2     | August 12, 2026  | Upgraded to SPFx 1.23.2 and migrated from Gulp to Heft |
| 1.1     | March 10, 2021   | Update comment  |
| 1.0     | January 29, 2021 | Initial release |

## Disclaimer

**THIS CODE IS PROVIDED _AS IS_ WITHOUT WARRANTY OF ANY KIND, EITHER EXPRESS OR IMPLIED, INCLUDING ANY IMPLIED WARRANTIES OF FITNESS FOR A PARTICULAR PURPOSE, MERCHANTABILITY, OR NON-INFRINGEMENT.**

---

## Minimal Path to Awesome

- Clone this repository
- Ensure that you are at the solution folder
- In the command-line, run:
  - `npm install`
  - `npm start` (to run the local workbench via `heft start`)
- To build and package for production:
  - `npm run build` (generates the `.sppkg` file in `sharepoint/solution`)

## Features

This extension illustrates the following concepts and features:

- **Dynamic Questions:** Pulls questions directly from a SharePoint list, making the survey easily maintainable without code changes.
- **URL Parameter Extraction:** Decrypts a Base64-encoded `encrypted` parameter from the URL to extract a Tracking ID.
- **Configurable Property Pane:** Allows site admins to select the target list, question list, and map specific fields directly from the SPFx Property Pane.
- **Modern SPFx Tooling:** Uses the latest `@rushstack/heft` build toolchain instead of `gulp`.

## References

- [Getting started with SharePoint Framework](https://docs.microsoft.com/en-us/sharepoint/dev/spfx/set-up-your-developer-tenant)
- [Microsoft 365 Patterns and Practices](https://aka.ms/m365pnp) - Guidance, tooling, samples and open-source controls for your Microsoft 365 development
