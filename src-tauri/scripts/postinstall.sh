#!/bin/sh
# Register the .ebb type installed at /usr/share/mime/packages/ebb.xml.
# The XML is inert until the database is rebuilt, and the .desktop entry only
# routes double-clicks once the desktop database knows about the app.
#
# Both steps are best-effort: a missing tool or a container without a MIME
# database costs the file association, never the installation.
set -e

if command -v update-mime-database >/dev/null 2>&1; then
    update-mime-database /usr/share/mime || true
fi

if command -v update-desktop-database >/dev/null 2>&1; then
    update-desktop-database /usr/share/applications || true
fi

exit 0
