/** Profile editing helpers, loaded with the Settings view rather than startup. */
import { providerRequestConfig } from './provider-profiles.js';

/** @param {Object<string, any>} entry */
export function syncActiveProviderProfile(entry) {
    const profile = entry.profiles?.find((item) => item.id === entry.activeProfileId);
    if (profile) profile.config = providerRequestConfig(entry);
}

/** Lazily migrate a single-key entry without changing its request settings.
 * @param {Object<string, any>} entry */
export function ensureProviderProfiles(entry) {
    if (!entry.profiles?.length) {
        entry.profiles = [{ id: 'default', name: 'Default', config: providerRequestConfig(entry) }];
        entry.activeProfileId = 'default';
    }
    return entry.profiles;
}

/** @param {Object<string, any>} entry @param {string} id */
export function selectProviderProfile(entry, id) {
    const profile = ensureProviderProfiles(entry).find((item) => item.id === id);
    if (!profile) return false;
    syncActiveProviderProfile(entry);
    Object.assign(entry, profile.config);
    entry.activeProfileId = id;
    return true;
}

/** A new profile inherits endpoint/model preferences, with an empty key.
 * @param {Object<string, any>} entry */
export function addProviderProfile(entry) {
    const profiles = ensureProviderProfiles(entry);
    syncActiveProviderProfile(entry);
    let number = 2;
    while (profiles.some((item) => item.id === `key-${number}`)) number += 1;
    const profile = {
        id: `key-${number}`, name: `API key ${number}`,
        config: { ...providerRequestConfig(entry), apiKey: '' },
    };
    profiles.push(profile);
    selectProviderProfile(entry, profile.id);
    return profile;
}

/** @param {Object<string, any>} entry */
export function removeProviderProfile(entry) {
    const profiles = ensureProviderProfiles(entry);
    if (profiles.length <= 1) return false;
    entry.profiles = profiles.filter((item) => item.id !== entry.activeProfileId);
    const next = entry.profiles[0];
    Object.assign(entry, next.config);
    entry.activeProfileId = next.id;
    return true;
}
