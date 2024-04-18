## Fun with Office Macros

This repository contains companion content that accompanied my 'Fun with Office Macros' BHIS webcast. The slides used for the presentation are contained in the root of this repository. The subfolders contain contents as described below:

- directpersistence - Example scripts to create restricted SSH users and macro content to use that SSH access to establish direct persistence.
- reconnaissance - Contains two macros for gathering useful information from an Active Directory client system under the context of a compromised user.
  - Reconnaissance_Simple.vba - This macro has most subroutines marked as private to provide simple operation of the underlying collection mechanism. The user is presented with the minimal number of macro subroutines to execute reconnaissance against a system or user.
  - Reconnaissance_Granular.vba - This macro has all subroutines marked public. This allows the operator to selectively execute various functionality at will rather than collecting all information from the environment in one shot.

When using the macros for document poisoning, I often remove unnecessary subroutines and focus on only the details that I need to collect.  For instance, I would independently run the BuildReconWorksheets, HideReconWorksheets, and UnHideReconWorksheets subroutines on the poisoned document to create the collection infrastructure. Then I would only include the functionality that makes sense in the context of the given scenario. Long running subroutines (Active Directory User, Group, and Computer collection)  may raise suspicion or frustrate users, as Ecxcel will be unusable until collection is complete. Subroutines that might not be OPSEC safe (Domain Trust Enumeration) may also need to be omitted, depending on the circumstances of execution.

Please consider expanding the suite of tooling that uses native Microsoft 365 product features.
