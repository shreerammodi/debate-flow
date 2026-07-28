import nextCoreWebVitals from "eslint-config-next/core-web-vitals";

const eslintConfig = [
    ...nextCoreWebVitals,
    // Nested git worktrees ship their own copy of the tree, including built
    // output that no ignore of theirs reaches from here, which otherwise
    // buries a root lint under hundreds of findings about generated code.
    {
        ignores: ["out/**", ".next/**", "src-tauri/**", "**/.worktrees/**", "**/.claude/**"],
    },
    {
        rules: {
            // These fire on guarded, keyed synchronizations (load-on-mount,
            // reset-on-close, interval ticks), not the cascading-render loop the
            // rule targets. Keep as a warning rather than blocking the build.
            "react-hooks/set-state-in-effect": "warn",
        },
    },
];

export default eslintConfig;
