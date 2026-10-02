(function () {
    // Same host the rest of Task Manager is deployed on (VPN/LAN-only by
    // design). This add-in intentionally does not try to work from outside
    // that network. Built from parts rather than a single literal string
    // since this file lives in a public repo (manifest.xml's AppDomains
    // entry still has to spell it out in full - that one's unavoidable,
    // Outlook/Office parse it directly) - keeps this one file at least out
    // of a plain-text search for the hostname.
    var API_BASE = 'https://' + ['tm-bms', 'duckdns', 'org'].join('.');
    var TOKEN_KEY = 'apiToken';

    var root = document.getElementById('root');

    function hasRoamingSettings() {
        return !!(Office.context && Office.context.roamingSettings);
    }

    function getToken() {
        if (!hasRoamingSettings()) return '';
        return Office.context.roamingSettings.get(TOKEN_KEY) || '';
    }

    function setToken(value, callback) {
        if (!hasRoamingSettings()) return;
        Office.context.roamingSettings.set(TOKEN_KEY, value);
        Office.context.roamingSettings.saveAsync(callback || function () {});
    }

    function clearToken(callback) {
        if (!hasRoamingSettings()) return;
        Office.context.roamingSettings.remove(TOKEN_KEY);
        Office.context.roamingSettings.saveAsync(callback || function () {});
    }

    function escapeHtml(value) {
        var div = document.createElement('div');
        div.textContent = value == null ? '' : String(value);
        return div.innerHTML;
    }

    // Outlook doesn't populate officeTheme.isDarkTheme (that only works on
    // other Office hosts) - darkness has to be derived from the actual
    // background color Outlook reports instead. Falls back to the OS/
    // browser's own color-scheme preference when officeTheme isn't
    // available at all (older clients, or hosts where this API is known to
    // be unreliable, e.g. classic Outlook desktop on Mac).
    function isDarkTheme() {
        try {
            var bg = Office.context && Office.context.officeTheme && Office.context.officeTheme.bodyBackgroundColor;
            if (bg) {
                var hex = bg.replace('#', '');
                if (hex.length === 3) {
                    hex = hex[0] + hex[0] + hex[1] + hex[1] + hex[2] + hex[2];
                }
                if (hex.length === 6) {
                    var r = parseInt(hex.substring(0, 2), 16);
                    var g = parseInt(hex.substring(2, 4), 16);
                    var b = parseInt(hex.substring(4, 6), 16);
                    // Standard relative-luminance weighting, not a plain
                    // average - matches how perceived brightness actually
                    // works (green reads brighter than blue at the same
                    // numeric value).
                    var luminance = (0.299 * r + 0.587 * g + 0.114 * b) / 255;
                    return luminance < 0.5;
                }
            }
        } catch (e) {
            // officeTheme not available in this host/version - fall through.
        }

        return !!(window.matchMedia && window.matchMedia('(prefers-color-scheme: dark)').matches);
    }

    function applyTheme() {
        document.body.classList.toggle('ta-dark', isDarkTheme());
    }

    // ── Token screen ─────────────────────────────────────────────────────

    function renderTokenScreen(message) {
        root.innerHTML =
            '<h1>BMS Task Manager</h1>' +
            '<p class="ta-subtitle">Vložte přístupový token, abyste mohli vytvářet úkoly přímo z Outlooku. Token najdete v Task Manageru v nabídce účtu &rarr; Outlook doplněk.</p>' +
            (message ? '<div class="ta-message ta-error">' + escapeHtml(message) + '</div>' : '') +
            '<div class="ta-field">' +
            '<label for="taTokenInput">Přístupový token</label>' +
            '<input type="text" id="taTokenInput" placeholder="1|xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx">' +
            '</div>' +
            '<button type="button" class="ta-primary" id="taSaveTokenBtn">Uložit a pokračovat</button>';

        document.getElementById('taSaveTokenBtn').addEventListener('click', function () {
            var value = document.getElementById('taTokenInput').value.trim();
            if (!value) return;

            setToken(value, function () {
                renderTaskForm();
            });
        });
    }

    // ── Task form ────────────────────────────────────────────────────────

    function currentSubject() {
        try {
            if (Office.context.mailbox && Office.context.mailbox.item && Office.context.mailbox.item.subject) {
                return Office.context.mailbox.item.subject;
            }
        } catch (e) {
            // No item in context (button opened with SupportsNoItemContext,
            // no email selected) - just start with a blank title.
        }
        return '';
    }

    function renderTaskForm(statusMessage, statusType) {
        root.innerHTML =
            '<h1>Nový úkol</h1>' +
            '<p class="ta-subtitle">Vytvořit úkol v BMS Task Manageru.</p>' +
            (statusMessage ? '<div class="ta-message ta-' + statusType + '">' + escapeHtml(statusMessage) + '</div>' : '') +
            '<div class="ta-field">' +
            '<label for="taTitle">Název *</label>' +
            '<input type="text" id="taTitle" value="' + escapeHtml(currentSubject()) + '">' +
            '</div>' +
            '<div class="ta-field">' +
            '<label for="taDescription">Popis</label>' +
            '<textarea id="taDescription"></textarea>' +
            '</div>' +
            '<div class="ta-row">' +
            '<div class="ta-field">' +
            '<label for="taDeadline">Termín</label>' +
            '<input type="date" id="taDeadline">' +
            '</div>' +
            '<div class="ta-field">' +
            '<label for="taPriority">Priorita</label>' +
            '<select id="taPriority">' +
            '<option value="Low">Nízká</option>' +
            '<option value="Medium" selected>Střední</option>' +
            '<option value="High">Vysoká</option>' +
            '</select>' +
            '</div>' +
            '</div>' +
            '<button type="button" class="ta-primary" id="taSaveTaskBtn">Vytvořit úkol</button>' +
            '<div class="ta-footer-link"><button type="button" class="ta-link" id="taChangeTokenBtn">Změnit token</button></div>';

        document.getElementById('taSaveTaskBtn').addEventListener('click', submitTask);
        document.getElementById('taChangeTokenBtn').addEventListener('click', function () {
            clearToken(function () {
                renderTokenScreen();
            });
        });

        document.getElementById('taTitle').focus();
    }

    function submitTask() {
        var title = document.getElementById('taTitle').value.trim();
        if (!title) {
            renderTaskForm('Název je povinný.', 'error');
            return;
        }

        var saveBtn = document.getElementById('taSaveTaskBtn');
        saveBtn.disabled = true;
        saveBtn.textContent = 'Ukládání…';

        var payload = { title: title };

        var description = document.getElementById('taDescription').value.trim();
        if (description) payload.description = description;

        var deadline = document.getElementById('taDeadline').value;
        if (deadline) payload.deadline = deadline;

        var priority = document.getElementById('taPriority').value;
        if (priority) payload.priority = priority;

        fetch(API_BASE + '/api/outlook-addin/tasks', {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json',
                Accept: 'application/json',
                Authorization: 'Bearer ' + getToken(),
            },
            body: JSON.stringify(payload),
        })
            .then(function (response) {
                if (response.status === 401) {
                    // Token missing/revoked/expired - clear it locally too,
                    // so the next open goes straight to re-pairing instead
                    // of repeating the same failed request.
                    clearToken(function () {
                        renderTokenScreen('Token již neplatí. Vygenerujte nový v Task Manageru a vložte ho znovu.');
                    });
                    return null;
                }

                return response.json().then(function (data) {
                    return { ok: response.ok, data: data };
                });
            })
            .then(function (result) {
                if (!result) return; // 401 path already handled above

                if (result.ok && result.data && result.data.success !== false) {
                    renderTaskForm('Úkol byl vytvořen.', 'success');
                } else {
                    var message = (result.data && result.data.message) || 'Úkol se nepodařilo vytvořit.';
                    renderTaskForm(message, 'error');
                }
            })
            .catch(function () {
                renderTaskForm(
                    'Nepodařilo se spojit s Task Managerem. Ujistěte se, že jste připojeni k firemní síti/VPN.',
                    'error'
                );
            });
    }

    // ── Entry point ──────────────────────────────────────────────────────

    Office.onReady(function () {
        try {
            applyTheme();

            // Best-effort: keep the theme in sync if the user switches
            // Outlook's theme while the pane is already open. Isolated in
            // its own try/catch - if this specific API isn't available on
            // some host/version, the pane should still render correctly
            // with whatever theme it detected at open time, not show an
            // error screen over a live-update nicety.
            try {
                if (Office.context && Office.context.officeTheme && typeof Office.addHandlerAsync === 'function') {
                    Office.addHandlerAsync(Office.EventType.OfficeThemeChanged, applyTheme);
                }
            } catch (themeHandlerError) {
                // Not supported on this host/version - already-applied
                // initial theme stands, nothing more to do.
            }
            if (window.matchMedia) {
                window.matchMedia('(prefers-color-scheme: dark)').addEventListener('change', applyTheme);
            }

            if (getToken()) {
                renderTaskForm();
            } else {
                renderTokenScreen();
            }
        } catch (e) {
            root.innerHTML =
                '<div class="ta-message ta-error">Nepodařilo se spustit doplněk v tomto zobrazení Outlooku.</div>';
        }
    });
})();
