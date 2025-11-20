# Customer Information Automated Updates Engine
## Executive Summary
This project establishes a cross-platform, fully automated data pipeline designed to process complex customer update requests. It integrates Python with GAS to ensure zero manual data handling during the reconciliation of sensitive information. 
The system successfully integrates complex business validation logic that was previously prone to human error, achieving complete automation while maintaining a critical human review checkpoint for final data verification.
## Core Technical Contributions
### End-to-End Automation & Integrated Workflow
* Full Automation: The entire workflow—from reading raw input requests to distributing finalized data and outputting check logs—is fully automated, significantly reducing the opportunity for human error in complex data entry and reconciliation tasks.
* Cross-Tool Integration: Successfully links the data-intensive capabilities of Python (Pandas/Pygsheets) for complex data mapping and ID parsing with the cloud-native automation of Google Apps Script (GAS) for final sheet operations and asynchronous triggering.
### Advanced Decision-Making and Complex Logic
* Conditional Data Routing: The Python logic uses ```tpr_updates``` and ```cbms_updates``` functions and implement two distinct decision scenarios (Restricted vs. Unrestricted), deciding whether to update the record immediately, insert a new row, or place the change into a pending queue.
* Data Integrity Checkpoints: The system enforces multiple checks (e.g., contract status, required fields) before performing high-risk updates, demonstrating dedication to data integrity.
### Output Validation and Human Review System
* Validation Output: The Python script generates detailed status logs (```tpr_check```, ```cbms_check```) for every processed Branch ID, summarizing the outcome (e.g., Y, H [Half-Finished], P [Pending], X [Not Found]).
* Final Synchronization: The GAS component synchronizes these check results back to a main sheet (```INPUT_SHEET.update_values```), allowing manual review and quality assurance to confirm system decisions before the data is utilized by downstream tools.
## Setup and Deployment
* Requires Node.js/npm and the Google Apps Script CLI (Clasp) for deployment.
* The Python engine securely retrieves these secrets at runtime using the ```google.colab.userdata```, preventing any hardcoding of credentials directly into the Notebook or script files.
