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

    // Real (non-signature/embedded-image) attachments on the current email -
    // a plain synchronous property in Outlook read mode, same as .subject
    // above, no async call needed just to list them.
    function currentAttachments() {
        try {
            var item = Office.context.mailbox && Office.context.mailbox.item;
            if (item && item.attachments) {
                return item.attachments.filter(function (a) { return !a.isInline; });
            }
        } catch (e) {
            // No item in context - same as currentSubject() above.
        }
        return [];
    }

    // Above this, getAttachmentContentAsync becomes slow/likely to fail in
    // practice (no hard documented ceiling from Office.js itself - this is
    // a UI hint, not an enforced limit, so an oversized file still gets
    // attempted, just with a warning shown up front).
    var ATTACHMENT_WARN_BYTES = 3 * 1024 * 1024;

    function getSelectedAttachments() {
        return Array.prototype.slice.call(document.querySelectorAll('.taAttachmentCheckbox:checked')).map(function (cb) {
            return {
                id: cb.getAttribute('data-att-id'),
                name: cb.getAttribute('data-att-name'),
                contentType: cb.getAttribute('data-att-content-type') || '',
            };
        });
    }

    function base64ToBlob(base64, contentType) {
        var byteChars = atob(base64);
        var byteNumbers = new Array(byteChars.length);
        for (var i = 0; i < byteChars.length; i++) {
            byteNumbers[i] = byteChars.charCodeAt(i);
        }
        return new Blob([new Uint8Array(byteNumbers)], { type: contentType || 'application/octet-stream' });
    }

    // Fetches one attachment's content from Outlook and uploads it straight
    // to the task's own attachments (not a comment) via the dedicated
    // Outlook-add-in endpoint - always resolves (never rejects) with
    // { ok, reason }, so Promise chains calling this never need their own
    // per-item catch, and failures carry an actual cause instead of just a
    // bare count (console.warn'd too, for anyone with devtools open on the
    // taskpane while diagnosing a report of "attachments don't work").
    function uploadOneAttachment(taskId, attachment) {
        return new Promise(function (resolve) {
            Office.context.mailbox.item.getAttachmentContentAsync(attachment.id, function (result) {
                if (result.status !== Office.AsyncResultStatus.Succeeded) {
                    var reason = 'nepodařilo se načíst přílohu z Outlooku' +
                        (result.error ? ' (' + result.error.message + ')' : '');
                    console.warn('Outlook add-in: attachment fetch failed for', attachment.name, result.error);
                    resolve({ ok: false, reason: reason });
                    return;
                }

                if (result.value.format !== Office.MailboxEnums.AttachmentContentFormat.Base64) {
                    // Cloud attachments (OneDrive links) and embedded-item
                    // attachments (forwarded emails) come back as a URL/EML
                    // instead of Base64 - not fetchable this way, skipped.
                    resolve({ ok: false, reason: 'cloudová/vložená příloha není podporována' });
                    return;
                }

                var blob = base64ToBlob(result.value.content, attachment.contentType);
                var formData = new FormData();
                formData.append('attachments[]', blob, attachment.name);

                fetch(API_BASE + '/api/outlook-addin/tasks/' + taskId + '/attachments', {
                    method: 'POST',
                    headers: {
                        Accept: 'application/json',
                        Authorization: 'Bearer ' + getToken(),
                    },
                    body: formData,
                })
                    .then(function (response) {
                        if (response.ok) {
                            resolve({ ok: true });
                            return;
                        }
                        return response.json().catch(function () { return {}; }).then(function (data) {
                            var reason = (data && data.message) || ('server vrátil chybu ' + response.status);
                            console.warn('Outlook add-in: attachment upload failed for', attachment.name, response.status, data);
                            resolve({ ok: false, reason: reason });
                        });
                    })
                    .catch(function (err) {
                        console.warn('Outlook add-in: attachment upload network error for', attachment.name, err);
                        resolve({ ok: false, reason: 'síťová chyba při nahrávání' });
                    });
            });
        });
    }

    // Uploads selected attachments one at a time (not Promise.all) - keeps
    // ordering deterministic and avoids firing several large uploads over
    // the VPN link at once.
    function uploadSelectedAttachments(taskId, attachments) {
        var okCount = 0;
        var failures = [];

        function next(index) {
            if (index >= attachments.length) {
                return Promise.resolve({ ok: okCount, failed: failures.length, failures: failures });
            }
            return uploadOneAttachment(taskId, attachments[index]).then(function (result) {
                if (result.ok) {
                    okCount++;
                } else {
                    failures.push(attachments[index].name + ': ' + result.reason);
                }
                return next(index + 1);
            });
        }

        return next(0);
    }

    // A native <input type="time"> renders its EMPTY state as a non-blank
    // "12:30"-looking segment placeholder on some Windows/Edge WebView2
    // builds instead of a clean "--:--" - the exact issue that already
    // forced a custom TmTimePicker in the main app. Using the same fix here
    // (two plain <select> elements) rather than a native time input avoids
    // the whole class of bug instead of re-hitting it in a second place.
    function timeOptionsHtml(range) {
        var html = '<option value="">--</option>';
        for (var i = 0; i < range; i++) {
            var value = i < 10 ? '0' + i : String(i);
            html += '<option value="' + value + '">' + value + '</option>';
        }
        return html;
    }

    // Only a real "HH:MM" when BOTH segments are picked - matches the old
    // single time-input's behavior where a half-filled time meant no time.
    function selectedDeadlineTime() {
        var hour = document.getElementById('taDeadlineHour').value;
        var minute = document.getElementById('taDeadlineMinute').value;
        return (hour && minute) ? (hour + ':' + minute) : '';
    }

    // Resolves to true/false, same convention as uploadOneAttachment -
    // never rejects, so the caller doesn't need its own catch.
    function setTaskDescription(taskId, description) {
        return fetch(API_BASE + '/api/outlook-addin/tasks/' + taskId + '/field', {
            method: 'PATCH',
            headers: {
                'Content-Type': 'application/json',
                Accept: 'application/json',
                Authorization: 'Bearer ' + getToken(),
            },
            body: JSON.stringify({ field: 'description', value: description }),
        })
            .then(function (response) { return response.ok; })
            .catch(function (err) {
                console.warn('Outlook add-in: setting description failed', err);
                return false;
            });
    }

    function loadAssignableUsers(selectEl) {
        fetch(API_BASE + '/api/outlook-addin/users', {
            headers: {
                Accept: 'application/json',
                Authorization: 'Bearer ' + getToken(),
            },
        })
            .then(function (response) { return response.ok ? response.json() : null; })
            .then(function (data) {
                if (!data || !data.users || !document.body.contains(selectEl)) return;
                data.users.forEach(function (user) {
                    var option = document.createElement('option');
                    option.value = user.id;
                    option.textContent = user.name;
                    selectEl.appendChild(option);
                });
            })
            .catch(function () {
                // Non-blocking - assignee picker just stays at "Já (výchozí)".
            });
    }

    function renderTaskForm(statusMessage, statusType) {
        var attachments = currentAttachments();
        var attachmentsFieldHtml = '';

        if (attachments.length) {
            attachmentsFieldHtml =
                '<div class="ta-field">' +
                '<label>Přílohy e-mailu</label>' +
                '<div class="ta-attachments-list">' +
                attachments.map(function (a) {
                    var warning = a.size > ATTACHMENT_WARN_BYTES
                        ? ' <span class="ta-att-warning">(velký soubor, nahrání může selhat)</span>'
                        : '';
                    return '<label class="ta-attachment-item">' +
                        '<input type="checkbox" class="taAttachmentCheckbox"' +
                        ' data-att-id="' + escapeHtml(a.id) + '"' +
                        ' data-att-name="' + escapeHtml(a.name) + '"' +
                        ' data-att-content-type="' + escapeHtml(a.contentType) + '">' +
                        '<span>' + escapeHtml(a.name) + warning + '</span>' +
                        '</label>';
                }).join('') +
                '</div>' +
                '</div>';
        }

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
            '<label class="ta-checkbox-field">' +
            '<input type="checkbox" id="taIncludeBody">' +
            '<span>Vložit text e-mailu do popisu</span>' +
            '</label>' +
            '<div class="ta-row">' +
            '<div class="ta-field">' +
            '<label for="taDeadline">Termín</label>' +
            '<input type="date" id="taDeadline">' +
            '</div>' +
            '<div class="ta-field">' +
            '<label for="taDeadlineHour">Čas</label>' +
            '<div class="ta-time-row">' +
            '<select id="taDeadlineHour">' + timeOptionsHtml(24) + '</select>' +
            '<span class="ta-time-sep">:</span>' +
            '<select id="taDeadlineMinute">' + timeOptionsHtml(60) + '</select>' +
            '</div>' +
            '</div>' +
            '</div>' +
            '<div class="ta-row">' +
            '<div class="ta-field">' +
            '<label for="taPriority">Priorita</label>' +
            '<select id="taPriority">' +
            '<option value="Low">Nízká</option>' +
            '<option value="Medium" selected>Střední</option>' +
            '<option value="High">Vysoká</option>' +
            '</select>' +
            '</div>' +
            '<div class="ta-field">' +
            '<label for="taAssignee">Přiřazeno</label>' +
            '<select id="taAssignee">' +
            '<option value="">Já (výchozí)</option>' +
            '</select>' +
            '</div>' +
            '</div>' +
            attachmentsFieldHtml +
            '<button type="button" class="ta-primary" id="taSaveTaskBtn">Vytvořit úkol</button>' +
            '<div class="ta-footer-link"><button type="button" class="ta-link" id="taChangeTokenBtn">Změnit token</button></div>';

        document.getElementById('taSaveTaskBtn').addEventListener('click', submitTask);
        document.getElementById('taChangeTokenBtn').addEventListener('click', function () {
            clearToken(function () {
                renderTokenScreen();
            });
        });

        loadAssignableUsers(document.getElementById('taAssignee'));

        document.getElementById('taTitle').focus();
    }

    // Only fetched when the user opts in (the checkbox is unchecked by
    // default) - item.body.getAsync is async, unlike .subject, so this has
    // to happen before the create-task request is built, not inline in the
    // payload construction below.
    function withDescription(baseDescription, includeBody, callback) {
        if (!includeBody) {
            callback(baseDescription);
            return;
        }

        try {
            Office.context.mailbox.item.body.getAsync(Office.CoercionType.Text, function (result) {
                var bodyText = result.status === Office.AsyncResultStatus.Succeeded ? result.value.trim() : '';
                callback(bodyText ? (baseDescription ? baseDescription + '\n\n---\n\n' + bodyText : bodyText) : baseDescription);
            });
        } catch (e) {
            callback(baseDescription);
        }
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

        var descriptionInput = document.getElementById('taDescription').value.trim();
        var includeBody = document.getElementById('taIncludeBody').checked;
        var selectedAttachments = getSelectedAttachments();

        withDescription(descriptionInput, includeBody, function (description) {
            // createNewTask has no description field (every task-creation
            // path in this app, including its own web form, adds the
            // description afterward via a separate field-update call, not
            // at creation time) - set it as a follow-up request below
            // instead of in this payload, which the backend would just
            // silently ignore.
            var payload = { title: title };

            var deadlineDate = document.getElementById('taDeadline').value;
            var deadlineTime = selectedDeadlineTime();
            if (deadlineDate) {
                payload.deadline = deadlineTime ? (deadlineDate + ' ' + deadlineTime + ':00') : deadlineDate;
            }

            var priority = document.getElementById('taPriority').value;
            if (priority) payload.priority = priority;

            var assignee = document.getElementById('taAssignee').value;
            if (assignee) payload.assigned_user_id = assignee;

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
                        var taskId = result.data.task && result.data.task.id;

                        var afterDescription = (taskId && description)
                            ? setTaskDescription(taskId, description)
                            : Promise.resolve(true);

                        afterDescription.then(function (descriptionOk) {
                            var baseMessage = 'Úkol byl vytvořen.' + (descriptionOk ? '' : ' Popis se nepodařilo uložit.');

                            if (taskId && selectedAttachments.length) {
                                uploadSelectedAttachments(taskId, selectedAttachments).then(function (summary) {
                                    var message = baseMessage;
                                    if (summary.failed) {
                                        message += ' ' + summary.ok + '/' + selectedAttachments.length + ' příloh se podařilo nahrát (' +
                                            summary.failures.join('; ') + ').';
                                        renderTaskForm(message, summary.ok ? 'success' : 'error');
                                    } else {
                                        renderTaskForm(message + ' Všechny přílohy byly nahrány.', 'success');
                                    }
                                });
                            } else {
                                renderTaskForm(baseMessage, descriptionOk ? 'success' : 'error');
                            }
                        });
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
