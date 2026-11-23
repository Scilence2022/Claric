import {
    normalizeWritingSettings, parseWritingDetection, sampleWritingText,
    writingDetectionMessages, WRITING_FIELDS, WRITING_OPTIONS,
} from '../lib/document-writing-settings.js';

export function createDocumentWritingController({
    identity, readText, getBackendConfig = () => ({}), classify, log, storage = globalThis.localStorage,
}) {
    const key = identity.documentEphemeral ? null : `claric.writing.${identity.documentId}`;
    const readStored = () => {
        try {
            const stored = key && storage.getItem(key);
            return stored ? normalizeWritingSettings(JSON.parse(stored)) : null;
        } catch {
            return null;
        }
    };
    let settings = readStored() || normalizeWritingSettings();
    let storedSettings = JSON.stringify(settings);
    let suggestion = null;
    let previousSettings = null;
    let lastAutomaticCheck = null;
    let pendingDetection = null;
    let detecting = false;
    let detectionFailed = false;
    let revision = 0;
    let status = '';
    const listeners = new Set();
    const snapshot = () => ({
        settings: normalizeWritingSettings(settings), detecting, detectionFailed, suggestion, canUndo: previousSettings !== null, status,
    });
    const notify = () => listeners.forEach((listener) => listener(snapshot()));

    function refresh() {
        const stored = readStored();
        if (!stored || JSON.stringify(stored) === storedSettings) return;
        settings = stored;
        storedSettings = JSON.stringify(stored);
        revision++;
        lastAutomaticCheck = null;
        suggestion = null;
        previousSettings = null;
        status = '';
        detectionFailed = false;
        notify();
    }

    function save(next, message = '') {
        settings = normalizeWritingSettings(next);
        revision++;
        status = message;
        if (key) {
            try {
                const serialized = JSON.stringify(settings);
                storage.setItem(key, serialized);
                storedSettings = serialized;
            } catch {
                status = [message, 'Settings could be kept only for this session.'].filter(Boolean).join(' ');
                log?.(status, 'warning');
            }
        }
        notify();
    }

    function applyDetection(detection, activityLog = log) {
        const entries = Object.entries(detection.values)
            .filter(([field, value]) => !settings.locked[field] && value !== settings.values[field]);
        if (!entries.length) {
            save(settings, 'No changes to unlocked settings were suggested.');
            return;
        }
        previousSettings = { values: settings.values, locked: settings.locked };
        const summary = `Updated ${entries.map(([field]) => field).join(', ')} from the document.`;
        save({ ...settings, values: { ...settings.values, ...Object.fromEntries(entries) } },
            [summary, detection.reason].filter(Boolean).join(' '));
        activityLog?.(`Document writing settings updated (${entries.map(([field, value]) => `${field}: ${WRITING_OPTIONS[field][value] || 'document context'}`).join(', ')}).`, 'info');
    }

    function clearTypeSuggestion() {
        if (!suggestion?.values.type) return;
        const { type: _obsoleteType, ...values } = suggestion.values;
        suggestion = Object.keys(values).length ? { ...suggestion, values } : null;
    }

    async function runDetection({ automatic, typeOnly, signal, onStart, log: activityLog = log }) {
        const startedAt = revision;
        const previousStatus = status;
        const updateType = automatic || (typeOnly && settings.mode === 'automatic');
        detecting = true;
        detectionFailed = false;
        if (!typeOnly) suggestion = null;
        status = 'Checking the document…';
        notify();
        const isCurrent = () => {
            refresh();
            return !signal?.aborted && revision === startedAt;
        };
        try {
            const sample = sampleWritingText(await readText());
            if (!isCurrent()) return;
            const backend = getBackendConfig();
            const model = JSON.stringify([backend.provider, backend.url, backend.apiPath, backend.model]);
            if (automatic && sample === lastAutomaticCheck?.sample && model === lastAutomaticCheck.model) {
                status = previousStatus;
                return;
            }
            let result = { values: {}, reason: '' };
            if (sample) {
                onStart?.();
                const reply = await classify(writingDetectionMessages(sample, { typeOnly }), signal, activityLog, backend);
                result = parseWritingDetection(reply, { typeOnly });
            }
            if (!isCurrent()) return;
            const detectedType = result.values.type || '';
            settings = { ...settings, detectedType };
            if (automatic) lastAutomaticCheck = { sample, model };
            if (typeOnly || !sample) clearTypeSuggestion();
            if (updateType && (sample || (settings.values.type && !settings.locked.type))) {
                applyDetection({ ...result, values: { ...result.values, type: detectedType } }, activityLog);
            } else if (!sample) {
                save(settings, 'Add text to the document to detect writing settings.');
            } else if (typeOnly) {
                save(settings, result.reason || (detectedType ? 'Document type refreshed.' : 'No clear document type detected.'));
            } else {
                const values = Object.fromEntries(Object.entries(result.values)
                    .filter(([field, value]) => !settings.locked[field] && value !== settings.values[field]));
                suggestion = Object.keys(values).length ? { ...result, values } : null;
                save(settings, suggestion ? (result.reason || 'Review the detected settings.') : 'No changes to unlocked settings were suggested.');
            }
        } catch (error) {
            if (isCurrent()) {
                status = 'Writing settings could not be detected.';
                detectionFailed = true;
                activityLog?.(`Writing settings detection failed: ${error.message}`, 'warning');
            }
        } finally {
            detecting = false;
            if (status === 'Checking the document…') status = '';
            notify();
        }
    }

    async function detect({ automatic = false, typeOnly = false, signal, ...options } = {}) {
        if (automatic) {
            // A turn must wait for a Refresh type or Detect settings already in progress.
            while (pendingDetection && !signal?.aborted) await pendingDetection;
        }
        refresh();
        if (pendingDetection || signal?.aborted) return;
        if (automatic) {
            if (settings.mode === 'manual') return;
            typeOnly = settings.mode === 'suggest';
        }
        if (!typeOnly && (settings.mode === 'manual' || WRITING_FIELDS.every((field) => settings.locked[field]))) return;
        pendingDetection = runDetection({ ...options, automatic, typeOnly, signal })
            .finally(() => { pendingDetection = null; });
        return pendingDetection;
    }
    return {
        get settings() {
            refresh();
            return normalizeWritingSettings(settings);
        },
        subscribe(listener) {
            refresh();
            listeners.add(listener);
            listener(snapshot());
        },
        useDetectedType() {
            refresh();
            const type = settings.detectedType;
            if (!type || (!settings.locked.type && type === settings.values.type)) return;
            previousSettings = { values: settings.values, locked: settings.locked };
            save({ ...settings, values: { ...settings.values, type }, locked: { ...settings.locked, type: false } },
                'Using the detected document type.');
        },
        update(change) {
            refresh();
            lastAutomaticCheck = null;
            suggestion = null;
            previousSettings = null;
            save({
                ...settings,
                mode: change.mode ?? settings.mode,
                values: { ...settings.values, ...change.values },
                locked: { ...settings.locked, ...change.locked },
            });
        },
        detect,
        applySuggestion() {
            refresh();
            if (!suggestion) return;
            const detection = suggestion;
            suggestion = null;
            applyDetection(detection);
        },
        undo() {
            refresh();
            if (!previousSettings) return;
            const previous = previousSettings;
            previousSettings = null;
            save({ ...settings, ...previous }, 'Previous writing settings restored.');
        },
    };
}
