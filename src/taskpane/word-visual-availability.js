import { screenCaptureState } from './word-screen-capture.js';

export function canReadWordRendering() {
    return typeof Office !== 'undefined' && ['Mac', 'PC', 'iOS'].includes(Office.context?.platform)
        && typeof Office.context?.document?.getFileAsync === 'function';
}

export function canReadWordVisuals() {
    return canReadWordRendering() || screenCaptureState().active;
}
