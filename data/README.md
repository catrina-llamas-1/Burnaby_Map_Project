# Input spreadsheet

Upload the worker roster spreadsheet here. Keep exactly **one** file in
this folder — `.xlsx`, `.xls`, and `.csv` all work, any filename.

To update the live map: open this folder on GitHub, **Add file → Upload
files**, and replace the existing spreadsheet (delete the old one first if
the filename changes). Pushing/merging that change automatically rebuilds
and redeploys the map — see the repo root README's "Deploying to Firebase
Hosting" section.

Required columns: **Worker's Postal**, **Lat**, **Long**, **Region**,
**Work Address - Line 1**, **City**, **Postal or ZIP code**. See
`generate_map.py`'s `CONFIG` block for exactly how each is used.

This spreadsheet can contain real employee postal codes and addresses —
only upload it to a private repository.
