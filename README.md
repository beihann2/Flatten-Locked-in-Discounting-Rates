# Flatten Locked-in Discounting Rates

This project aims to flatten discounting locked-in rates for every contract group in accordance with IFRS 17 Standards. 

### Objective
The primary aim is to reduce the asset-liability duration gap in future financial reporting periods while ensuring full alignment with IFRS 17 requirements. Specifically, this project seeks to:
1. Stabilize future insurance expenses to support long-term financial health.
2. Narrow the duration gap between assets and liabilities.
3. Automate data processing to improve efficiency and accuracy.
4. Implement resource-efficient solutions that maintain system integrity.


### Features
1. Calculation of flattened discount rates using Newton’s method to goal-seek a flat rate for each contract group.
2. Supports both insurance contracts and reinsurance contracts.
3. Goal-seeking algorithm ensures that the present value of future fulfillment cash flows and the risk adjustment—factoring in the loss-absorbing effect—remains consistent with current liabilities.
4. Cash flow projections are derived from actuarial models covering up to 106 years, in line with IFRS 17 standards. The project utilizes FIS Prophet for actuarial modeling.

## Data Files
The repository includes the following Excel file:
- CFS.xlsm: Contains sample data of insurance contracts used for calculations with VBA codes embedded.
- CFS-Reins.xlsm: Contains sample data of reinsurance contracts used for calculations with VBA codes embedded.


### Getting Started
1. Clone the repository:
   git clone https://github.com/username/Flatten-Locked-in-Discounting-Rates.git
2. Modify the relevant code and file paths in your local environment.
3. Run the calculations to obtain the flattened discount rates.

