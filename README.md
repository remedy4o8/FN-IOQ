Freenome IOQ Export Assistant

A PowerShell WPF GUI tool for lab technicians to collect machine compliance data and export it as a CSV, either directly to a NetApp share or locally to the Desktop as a backup.

What It Does

Collects system info (PC name, model, serial, OS version, domain, security settings)
Checks installed software versions for a defined list of required apps
Tags the export with the technician's name and the current date
Uploads the CSV directly to a NetApp network share, or saves it locally if the upload fails
