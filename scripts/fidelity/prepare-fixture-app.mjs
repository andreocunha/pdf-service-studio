// Builds an isolated, local-only app. Real documents stay outside the repository.
import fs from 'node:fs/promises';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
const here=path.dirname(fileURLToPath(import.meta.url));
const [source,target]=process.argv.slice(2).map(p=>path.resolve(p));
if(!source||!target||source===target)throw new Error('Usage: node prepare-fixture-app.mjs APP_SOURCE NEW_TEMP_APP');
await fs.mkdir(target); // Never overwrite a working application.
for(const name of ['app','public','docs','scripts'])await fs.cp(path.join(source,name),path.join(target,name),{recursive:true});
for(const name of ['package.json','package-lock.json','tsconfig.json','next.config.ts','next-env.d.ts','proxy.ts','postcss.config.mjs','global.d.ts','instrumentation.ts','.env.local']){
 try{await fs.copyFile(path.join(source,name),path.join(target,name));}catch(e){if(e.code!=='ENOENT')throw e;}
}
await fs.symlink(path.join(source,'node_modules'),path.join(target,'node_modules'),'dir');
const config=path.join(target,'next.config.ts');
await fs.writeFile(config,(await fs.readFile(config,'utf8')).replace('const nextConfig: NextConfig = {','const nextConfig: NextConfig = {\n  turbopack: { root: "/" },'));
const route=path.join(target,'app/render-pdf/validation/[id]');await fs.mkdir(route,{recursive:true});
for(const name of ['page.tsx','validation-client.tsx'])await fs.copyFile(path.join(here,name+'.fixture'),path.join(route,name));
console.log('Prepared isolated fixture app: '+target);
console.log('Start it with FIDELITY_FIXTURE_DIR pointing at local UUID.json snapshots. No document edits are saved.');
