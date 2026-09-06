/** Cooperative named resource locks for task execution. */
export function createResourceLocks() {
    const owners = new Map();
    return {
        canAcquire(resources = [], owner) { return resources.every((resource) => !owners.has(resource) || owners.get(resource) === owner); },
        acquire(resources = [], owner) { if (!this.canAcquire(resources, owner)) return false; resources.forEach((resource) => owners.set(resource, owner)); return true; },
        release(resources = [], owner) { resources.forEach((resource) => { if (owners.get(resource) === owner) owners.delete(resource); }); },
        snapshot() { return new Map(owners); },
    };
}
