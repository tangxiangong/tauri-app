default: dev

tauri *args:
    deno task tauri {{args}}

# Tauri development mode
dev:
    deno task tauri dev

# Build Tauri app for production
build:
    deno task tauri build

# Generate Tauri icons
icon *args:
    deno task tauri icon {{args}} -o src-tauri/icons

web *args:
    deno task vite {{args}}

web-dev:
    deno task dev --open
