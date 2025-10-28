// Ambient helper for third-party typings that reference ReadonlySetLike
// MobX's types may reference ReadonlySetLike which isn't present in some TS lib setups.
// Provide a compatible alias so the project can compile.
type ReadonlySetLike<T> = ReadonlySet<T> | Set<T> | T[];
