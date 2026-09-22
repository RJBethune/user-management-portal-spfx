# Package identity and filenames

Account Management uses the same original solution, feature and component IDs in DEV and PROD. Channel and version stay in release metadata and settings attribution. Toolkit 1.5.3 rejects split identity records before staging and checks the emitted names, IDs, component counts, asset references and all-sites eligibility.

Installing either channel updates the original application in the selected catalog. Use a separate test catalog or tenant for DEV isolation. Existing registrations created by older DEV packages are not removed automatically. Inventory their page usage and preserve page configuration before any separately authorized cleanup. This packaging change does not migrate lists, permissions or data.

## Upload filenames

Both channel folders use `account-management.sppkg`. You may rename the outer `.sppkg` file or a downloaded copy while retaining its extension. The package bytes carry the identity; renaming cannot change its IDs, names, permissions or asset references. Toolkit 1.5.3 and the updated Portfolio recognize a renamed file only when exactly one package in the release folder matches the original SHA-256. They preserve the original manifest and checksum inventory. Modified packages, ambiguous copies and renamed runtime assets fail verification. Static index links reflect filenames at generation time; use Portfolio for later renames.

## Original identities

```json
{
  "solutionId": "ee656af9-b263-4758-b163-a56663bc6cbd",
  "featureId": "0ce47f4e-35ca-4b99-afb7-ce1b2703ccaf",
  "componentId": "762df7fa-0fd0-4019-b227-d184ab4abc67"
}
```

## Obsolete DEV identities

Recovery references only; never restore these to active delivery configuration.

```json
{
  "solutionId": "cc3432be-8ecc-4be6-a9ee-1652f687c9d7",
  "featureId": "21ecc7d5-b080-4730-86bf-ced482879e12",
  "componentId": "def2a872-bf6d-4766-bf9a-ffb05bd35fe6"
}
```
