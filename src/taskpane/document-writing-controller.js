import {
    normalizeWritingSettings, parseWritingDetection, sampleWritingText,
    writingDetectionMessages, WRITING_FIELDS, WRITING_OPTIONS,
} from '../lib/document-writing-settings.js';

export function createDocumentWritingController({ identity, readText, classify, log, storage = globalThis.localStorage }) {
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
    let previousValues = null;
    let lastSample = null;
    let detecting = false;
    let revision = 0;
    let status = '';
    let lastDetectionStatus = '';
    const listeners = new Set();
    const snapshot = () => ({
        settings: normalizeWritingSettings(settings), detecting, suggestion, canUndo: previousValues !== null, status,
    });
    const notify = () => listeners.forEach((listener) => listener(snapshot()));

    function refresh() {
        const stored = readStored();
        if (!stored || JSON.stringify(stored) === storedSettings) return;
        settings = stored;
        storedSettings = JSON.stringify(stored);
        revision++;
        lastSample = null;
        suggestion = null;
        previousValues = null;
        status = '';
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

    function applyDetection(detection) {
        const entries = Object.entries(detection.values)
            .filter(([field, value]) => !settings.locked[field] && value !== settings.values[field]);
        suggestion = null;
        if (!entries.length) {
            status = 'No changes to unlocked settings were suggested.';
            notify();
            return;
        }
        previousValues = { ...settings.values };
        const summary = `Updated ${entries.map(([field]) => field).join(', ')} from the document.`;
        save({ ...settings, values: { ...settings.values, ...Object.fromEntries(entries) } },
            [summary, detection.reason].filter(Boolean).join(' '));
        log?.(`Document writing settings updated (${entries.map(([field, value]) => `${field}: ${WRITING_OPTIONS[field][value]}`).join(', ')}).`, 'info');
    }

    async function detect({ automatic = false, signal, onStart } = {}) {
        refresh();
        if (detecting || signal?.aborted || settings.mode === 'manual'
            || (automatic && settings.mode !== 'automatic')
            || WRITING_FIELDS.every((field) => settings.locked[field])) return;
        const startedAt = revision;
        detecting = true;
        suggestion = null;
        status = 'Checking the document…';
        notify();
        const isCurrent = () => {
            refresh();
            return !signal?.aborted && revision === startedAt;
        };
        try {
            const sample = sampleWritingText(await readText());
            if (!isCurrent()) return;
            if (!sample || (automatic && sample === lastSample)) {
                status = sample ? lastDetectionStatus : 'Add text to the document to detect writing settings.';
                return;
            }
            onStart?.();
            const result = parseWritingDetection(await classify(writingDetectionMessages(sample), signal));
            if (!isCurrent()) return;
            if (automatic) {
                lastSample = sample;
                applyDetection(result);
                lastDetectionStatus = status;
            } else {
                const values = Object.fromEntries(Object.entries(result.values)
                    .filter(([field, value]) => !settings.locked[field] && value !== settings.values[field]));
                suggestion = Object.keys(values).length ? { ...result, values } : null;
                status = suggestion ? (result.reason || 'Review the detected settings.') : 'No changes to unlocked settings were suggested.';
            }
        } catch (error) {
            if (isCurrent()) {
                status = 'Writing settings could not be detected.';
                log?.(`Writing settings detection failed: ${error.message}`, 'warning');
            }
        } finally {
            detecting = false;
            if (status === 'Checking the document…') status = '';
            notify();
        }
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
        update(change) {
            refresh();
            lastSample = null;
            suggestion = null;
            previousValues = null;
            save({
                mode: change.mode ?? settings.mode,
                values: { ...settings.values, ...change.values },
                locked: { ...settings.locked, ...change.locked },
            });
        },
        detect,
        applySuggestion() {
            refresh();
            if (suggestion) applyDetection(suggestion);
        },
        undo() {
            refresh();
            if (!previousValues) return;
            const values = previousValues;
            previousValues = null;
            lastDetectionStatus = 'Previous writing settings restored.';
            save({ ...settings, values }, lastDetectionStatus);
        },
    };
}
