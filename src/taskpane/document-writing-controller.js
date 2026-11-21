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
    let previousSettings = null;
    let lastSample = null;
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
        lastSample = null;
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
        activityLog?.(`Document writing settings updated (${entries.map(([field, value]) => `${field}: ${WRITING_OPTIONS[field][value]}`).join(', ')}).`, 'info');
    }

    function clearTypeSuggestion() {
        if (!suggestion?.values.type) return;
        const { type: _obsoleteType, ...values } = suggestion.values;
        suggestion = Object.keys(values).length ? { ...suggestion, values } : null;
    }

    async function detect({ automatic = false, typeOnly = false, signal, onStart, log: activityLog = log } = {}) {
        refresh();
        if (detecting || signal?.aborted || (automatic && settings.mode !== 'automatic')) return;
        if (!typeOnly && (settings.mode === 'manual' || WRITING_FIELDS.every((field) => settings.locked[field]))) return;
        const startedAt = revision;
        const previousStatus = status;
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
            if (!sample) {
                clearTypeSuggestion();
                save({ ...settings, detectedType: '' }, 'Add text to the document to detect writing settings.');
                return;
            }
            if (automatic && sample === lastSample) {
                status = previousStatus;
                return;
            }
            onStart?.();
            const result = parseWritingDetection(await classify(writingDetectionMessages(sample, { typeOnly }), signal, activityLog), { typeOnly });
            if (!isCurrent()) return;
            settings = { ...settings, detectedType: result.values.type || '' };
            if (automatic) {
                lastSample = sample;
                applyDetection(result, activityLog);
            } else if (typeOnly) {
                clearTypeSuggestion();
                if (settings.mode === 'automatic' && !settings.locked.type && result.values.type) {
                    applyDetection(result, activityLog);
                } else {
                    save(settings, result.reason || (settings.detectedType ? 'Document type refreshed.' : 'No clear document type detected.'));
                }
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
            lastSample = null;
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
