# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Architecture

This is a Tauri desktop application using:
- **Frontend**: SvelteKit 5 with TypeScript, TailwindCSS v4, and DaisyUI
- **Backend**: Rust with Tauri 2.x framework
- **Build System**: Vite for frontend bundling, Cargo for Rust compilation
- **Runtime**: Deno for task running (see Justfile and tauri.conf.json)

### Key Architecture Points

- **SPA Mode**: Uses `@sveltejs/adapter-static` with fallback to `index.html` since Tauri doesn't support SSR
- **Tauri Commands**: Rust functions exposed to frontend via `#[tauri::command]` macro (see `src-tauri/src/lib.rs`)
- **Frontend-Backend Communication**: TypeScript wrappers in `src/lib/command/` call Rust functions via `@tauri-apps/api/core.invoke()`
- **State Management**: Svelte 5 runes (`$state`) for reactive state
- **Styling**: TailwindCSS v4 with DaisyUI component library

## Development Commands

### Frontend Development
```bash
npm run dev          # Start Vite dev server (frontend only)
npm run build        # Build frontend for production
npm run preview      # Preview production build
```

### Type Checking & Linting  
```bash
npm run check        # Run svelte-check with TypeScript
npm run check:watch  # Run svelte-check in watch mode
```

### Tauri Development
```bash
npm run tauri dev    # Start Tauri app in development mode
npm run tauri build  # Build Tauri app for production
```

### Alternative Task Runner (Deno)
```bash
just dev            # Equivalent to deno task tauri dev (uses Justfile)
```

## Project Structure

```
src/
├── routes/          # SvelteKit pages (main app entry in +page.svelte)
├── components/      # Reusable Svelte components
├── lib/
│   ├── command/     # TypeScript wrappers for Tauri commands
│   └── shared.svelte.ts # Shared reactive state
└── app.html         # HTML template

src-tauri/
├── src/
│   ├── lib.rs       # Tauri app setup and command handlers
│   └── main.rs      # Entry point
├── Cargo.toml       # Rust dependencies
└── tauri.conf.json  # Tauri configuration
```

## Adding New Tauri Commands

1. Add Rust function with `#[tauri::command]` in `src-tauri/src/lib.rs`
2. Register command in `tauri::generate_handler![]` macro
3. Create TypeScript wrapper in `src/lib/command/[name].ts` using `invoke()`
4. Import and use in Svelte components

## Configuration Notes

- **Development Server**: Fixed port 1420 (required by Tauri)
- **TypeScript**: Strict mode enabled with `noImplicitAny`
- **Bundle Target**: Static files for desktop distribution
- **App Identifier**: `com.xiaoyu.tauri-app`