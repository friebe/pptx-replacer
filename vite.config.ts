import { defineConfig } from 'vite';
import { builtinModules } from 'module';

export default defineConfig({
    build: {
        lib: {
            entry: './src/main.ts',
            formats: ['es'],
            fileName: () => 'processPPTX.js'
        },
        rollupOptions: {
            external: [
                ...builtinModules,
                'pizzip',
                'fs'
            ]
        },
        outDir: 'dist'
    }
});
