// Gắn biểu tượng + thông tin phiên bản vào ExcelLike.exe bằng resedit (thuần JS, không cần Wine khi đóng gói trên Linux)
const path = require('node:path');
const fs = require('node:fs');

exports.default = async function afterPack(ctx) {
  if (ctx.electronPlatformName !== 'win32') return;
  const ResEdit = await import('resedit');
  const { NtExecutable, NtExecutableResource, Resource, Data } = ResEdit.default ?? ResEdit;
  const pkg = ctx.packager.appInfo;
  const exe = path.join(ctx.appOutDir, `${pkg.productFilename}.exe`);
  const data = fs.readFileSync(exe);
  const ne = NtExecutable.from(data, { ignoreCert: true });
  const res = NtExecutableResource.from(ne);
  const ico = Data.IconFile.from(fs.readFileSync(path.join(__dirname, 'icon.ico')));
  Resource.IconGroupEntry.replaceIconsForResource(res.entries, 1, 1033, ico.icons.map((i) => i.data));
  const [vi] = Resource.VersionInfo.fromEntries(res.entries);
  const ver = pkg.version.split('.').map(Number);
  while (ver.length < 4) ver.push(0);
  vi.setFileVersion(...ver);
  vi.setProductVersion(...ver);
  const lang = { lang: 1033, codepage: 1200 };
  vi.setStringValues(lang, {
    FileDescription: 'ExcelLike – bảng tính',
    ProductName: 'ExcelLike',
    CompanyName: 'LAVIPCO',
    LegalCopyright: '© 2026 Mai Vũ Lâm',
    OriginalFilename: `${pkg.productFilename}.exe`,
    InternalName: 'ExcelLike',
    FileVersion: pkg.version,
    ProductVersion: pkg.version,
  });
  vi.outputToResourceEntries(res.entries);
  res.outputResource(ne);
  fs.writeFileSync(exe, Buffer.from(ne.generate()));
  console.log('  • đã gắn biểu tượng và thông tin phiên bản vào', path.basename(exe));
};
