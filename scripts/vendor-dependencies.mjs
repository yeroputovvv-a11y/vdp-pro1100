import { cp, mkdir, rm } from 'node:fs/promises';
import { join } from 'node:path';

const root = new URL('..', import.meta.url).pathname;
const nm = join(root, 'node_modules');
const vendor = join(root, 'vendor');

await rm(vendor, { recursive: true, force: true });
await mkdir(join(vendor, 'xlsx'), { recursive: true });
await mkdir(join(vendor, 'pdfjs'), { recursive: true });
await mkdir(join(vendor, 'jszip'), { recursive: true });
await mkdir(join(vendor, 'mupdf'), { recursive: true });
await mkdir(join(vendor, 'pdflib'), { recursive: true });
await mkdir(join(vendor, 'fonts'), { recursive: true });

await cp(join(nm, 'xlsx-js-style', 'dist', 'xlsx.bundle.js'), join(vendor, 'xlsx', 'xlsx.bundle.js'));
await cp(join(nm, 'pdfjs-dist', 'build', 'pdf.min.js'), join(vendor, 'pdfjs', 'pdf.min.js'));
await cp(join(nm, 'pdfjs-dist', 'build', 'pdf.worker.min.js'), join(vendor, 'pdfjs', 'pdf.worker.min.js'));
await cp(join(nm, 'jszip', 'dist', 'jszip.min.js'), join(vendor, 'jszip', 'jszip.min.js'));
await cp(join(nm, 'mupdf', 'dist'), join(vendor, 'mupdf'), { recursive: true });
await cp(join(nm, 'pdf-lib', 'dist', 'pdf-lib.min.js'), join(vendor, 'pdflib', 'pdf-lib.min.js'));
await cp(join(nm, 'dejavu-fonts-ttf', 'ttf', 'DejaVuSans.ttf'), join(vendor, 'fonts', 'DejaVuSans.ttf'));

console.log('Vendor assets created in ./vendor');
