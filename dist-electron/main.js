import Ht, { app as $e, ipcMain as me, dialog as Pn, shell as Fl, BrowserWindow as xl } from "electron";
import lt from "node:path";
import { fileURLToPath as sd } from "node:url";
import { createRequire as ld } from "node:module";
import Ce from "node:fs";
import Ct from "fs";
import cd from "constants";
import Zn from "stream";
import Po from "util";
import Ul from "assert";
import ee from "path";
import ei from "child_process";
import kl from "events";
import er from "crypto";
import Ml from "tty";
import ti from "os";
import bt from "url";
import Bl from "zlib";
import ud from "http";
function Pi(e = /* @__PURE__ */ new Date()) {
  return `${e.getFullYear()}-${String(e.getMonth() + 1).padStart(2, "0")}-${String(e.getDate()).padStart(2, "0")}`;
}
function fd(e) {
  return e ? JSON.parse(e) : [];
}
function La(e, t) {
  const n = fd(e == null ? void 0 : e.rent_history), r = Number(t.rent_amount);
  if (!Number.isFinite(r) || r < 0) throw new Error("Kira tutarı sıfır veya pozitif olmalıdır.");
  if (e && Number(e.rent_amount) === r && !t.rent_effective_date) return JSON.stringify(n);
  const i = t.rent_effective_date || (e ? "" : t.contract_start_date || Pi());
  if (!/^\d{4}-\d{2}-\d{2}$/.test(i) || !Number.isFinite(Date.parse(i)) || new Date(i).toISOString().slice(0, 10) !== i)
    throw new Error("Geçerli bir sözleşme güncelleme tarihi girin.");
  if (n.length && i <= n[n.length - 1].effective_date)
    throw new Error("Yeni tarih son sözleşme tarihinden sonra olmalıdır.");
  return !n.length && e && Pi() < i && n.push({ effective_date: Pi(), amount: Number(e.rent_amount) }), n.push({ effective_date: i, amount: r }), JSON.stringify(n);
}
const Hl = (e) => e.replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\uFFFE\uFFFF]/g, "").replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&apos;");
function Fa(e) {
  let t = "";
  for (let n = e + 1; n > 0; n = Math.floor((n - 1) / 26))
    t = String.fromCharCode(65 + (n - 1) % 26) + t;
  return t;
}
function dd(e) {
  if (typeof e != "string" || !/^\d{4}-\d{2}-\d{2}$/.test(e)) return null;
  const t = Date.parse(e + "T00:00:00Z");
  return !Number.isFinite(t) || new Date(t).toISOString().slice(0, 10) !== e ? null : t / 864e5 + 25569;
}
function hd(e) {
  if (!e.headers.length || e.headers.length > 16384 || e.rows.length > 1048575)
    throw new Error("Kayıt sayısı Excel sayfa sınırını aşıyor.");
  const t = `${Fa(e.headers.length - 1)}${e.rows.length + 1}`, n = [e.headers, ...e.rows].map((i, o) => {
    const a = i.map((s, c) => {
      var h, g;
      const p = `${Fa(c)}${o + 1}`, l = o > 0 && ((h = e.dateColumns) != null && h.includes(c)) ? dd(s) : null;
      if (l !== null) return `<c r="${p}" s="3"><v>${l}</v></c>`;
      if (typeof s == "number" && Number.isFinite(s))
        return `<c r="${p}" s="${(g = e.moneyColumns) != null && g.includes(c) ? 2 : 0}"><v>${s}</v></c>`;
      const f = String(s ?? "");
      if (f.length > 32767) throw new Error("Bir hücredeki metin Excel sınırı olan 32767 karakteri aşıyor.");
      return `<c r="${p}" s="${o === 0 ? 1 : 0}" t="inlineStr"><is><t xml:space="preserve">${Hl(f)}</t></is></c>`;
    }).join("");
    return `<row r="${o + 1}"${o === 0 ? ' ht="28" customHeight="1"' : ""}>${a}</row>`;
  }).join(""), r = e.headers.map(
    (i, o) => `<col min="${o + 1}" max="${o + 1}" width="${/Not|Açıklama|Mülk|Hesap/.test(i) ? 32 : 20}" customWidth="1"/>`
  ).join("");
  return `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><dimension ref="A1:${t}"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><cols>${r}</cols><sheetData>${n}</sheetData><autoFilter ref="A1:${t}"/><pageSetup orientation="landscape" paperSize="9"/></worksheet>`;
}
function pd(e) {
  let t = 4294967295;
  for (const n of e) {
    t ^= n;
    for (let r = 0; r < 8; r++) t = t >>> 1 ^ (t & 1 ? 3988292384 : 0);
  }
  return (t ^ 4294967295) >>> 0;
}
function md(e) {
  const t = [], n = [];
  let r = 0;
  for (const a of e) {
    const s = Buffer.from(a.name), c = Buffer.from(a.content, "utf8"), p = pd(c), l = Buffer.alloc(30);
    l.writeUInt32LE(67324752, 0), l.writeUInt16LE(20, 4), l.writeUInt16LE(2048, 6), l.writeUInt16LE(33, 12), l.writeUInt32LE(p, 14), l.writeUInt32LE(c.length, 18), l.writeUInt32LE(c.length, 22), l.writeUInt16LE(s.length, 26);
    const f = Buffer.alloc(46);
    f.writeUInt32LE(33639248, 0), f.writeUInt16LE(20, 4), f.writeUInt16LE(20, 6), f.writeUInt16LE(2048, 8), f.writeUInt16LE(33, 14), f.writeUInt32LE(p, 16), f.writeUInt32LE(c.length, 20), f.writeUInt32LE(c.length, 24), f.writeUInt16LE(s.length, 28), f.writeUInt32LE(r, 42), t.push(l, s, c), n.push(f, s), r += l.length + s.length + c.length;
  }
  const i = Buffer.concat(n), o = Buffer.alloc(22);
  return o.writeUInt32LE(101010256, 0), o.writeUInt16LE(e.length, 8), o.writeUInt16LE(e.length, 10), o.writeUInt32LE(i.length, 12), o.writeUInt32LE(r, 16), Buffer.concat([...t, i, o]);
}
function gd(e) {
  const t = "http://schemas.openxmlformats.org", n = [
    { name: "[Content_Types].xml", content: `<Types xmlns="${t}/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>${e.map((r, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join("")}</Types>` },
    { name: "_rels/.rels", content: `<Relationships xmlns="${t}/package/2006/relationships"><Relationship Id="rId1" Type="${t}/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>` },
    { name: "xl/workbook.xml", content: `<workbook xmlns="${t}/spreadsheetml/2006/main" xmlns:r="${t}/officeDocument/2006/relationships"><sheets>${e.map((r, i) => `<sheet name="${Hl(r.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join("")}</sheets></workbook>` },
    { name: "xl/_rels/workbook.xml.rels", content: `<Relationships xmlns="${t}/package/2006/relationships">${e.map((r, i) => `<Relationship Id="rId${i + 1}" Type="${t}/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join("")}<Relationship Id="styles" Type="${t}/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>` },
    { name: "xl/styles.xml", content: `<styleSheet xmlns="${t}/spreadsheetml/2006/main"><numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.00 &quot;₺&quot;"/><numFmt numFmtId="165" formatCode="dd.mm.yyyy"/></numFmts><fonts count="2"><font><sz val="11"/><name val="Calibri"/></font><font><b/><color rgb="FFFFFFFF"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF1E3A5F"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border/></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="4"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1"/><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/><xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>` },
    ...e.map((r, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, content: hd(r) }))
  ];
  return md(n);
}
var Oe = typeof globalThis < "u" ? globalThis : typeof window < "u" ? window : typeof global < "u" ? global : typeof self < "u" ? self : {}, ze = {}, Gt = {}, De = {};
De.fromCallback = function(e) {
  return Object.defineProperty(function(...t) {
    if (typeof t[t.length - 1] == "function") e.apply(this, t);
    else
      return new Promise((n, r) => {
        t.push((i, o) => i != null ? r(i) : n(o)), e.apply(this, t);
      });
  }, "name", { value: e.name });
};
De.fromPromise = function(e) {
  return Object.defineProperty(function(...t) {
    const n = t[t.length - 1];
    if (typeof n != "function") return e.apply(this, t);
    t.pop(), e.apply(this, t).then((r) => n(null, r), n);
  }, "name", { value: e.name });
};
var pt = cd, Ed = process.cwd, kr = null, yd = process.env.GRACEFUL_FS_PLATFORM || process.platform;
process.cwd = function() {
  return kr || (kr = Ed.call(process)), kr;
};
try {
  process.cwd();
} catch {
}
if (typeof process.chdir == "function") {
  var xa = process.chdir;
  process.chdir = function(e) {
    kr = null, xa.call(process, e);
  }, Object.setPrototypeOf && Object.setPrototypeOf(process.chdir, xa);
}
var _d = wd;
function wd(e) {
  pt.hasOwnProperty("O_SYMLINK") && process.version.match(/^v0\.6\.[0-2]|^v0\.5\./) && t(e), e.lutimes || n(e), e.chown = o(e.chown), e.fchown = o(e.fchown), e.lchown = o(e.lchown), e.chmod = r(e.chmod), e.fchmod = r(e.fchmod), e.lchmod = r(e.lchmod), e.chownSync = a(e.chownSync), e.fchownSync = a(e.fchownSync), e.lchownSync = a(e.lchownSync), e.chmodSync = i(e.chmodSync), e.fchmodSync = i(e.fchmodSync), e.lchmodSync = i(e.lchmodSync), e.stat = s(e.stat), e.fstat = s(e.fstat), e.lstat = s(e.lstat), e.statSync = c(e.statSync), e.fstatSync = c(e.fstatSync), e.lstatSync = c(e.lstatSync), e.chmod && !e.lchmod && (e.lchmod = function(l, f, h) {
    h && process.nextTick(h);
  }, e.lchmodSync = function() {
  }), e.chown && !e.lchown && (e.lchown = function(l, f, h, g) {
    g && process.nextTick(g);
  }, e.lchownSync = function() {
  }), yd === "win32" && (e.rename = typeof e.rename != "function" ? e.rename : function(l) {
    function f(h, g, _) {
      var E = Date.now(), T = 0;
      l(h, g, function A(R) {
        if (R && (R.code === "EACCES" || R.code === "EPERM" || R.code === "EBUSY") && Date.now() - E < 6e4) {
          setTimeout(function() {
            e.stat(g, function(P, B) {
              P && P.code === "ENOENT" ? l(h, g, A) : _(R);
            });
          }, T), T < 100 && (T += 10);
          return;
        }
        _ && _(R);
      });
    }
    return Object.setPrototypeOf && Object.setPrototypeOf(f, l), f;
  }(e.rename)), e.read = typeof e.read != "function" ? e.read : function(l) {
    function f(h, g, _, E, T, A) {
      var R;
      if (A && typeof A == "function") {
        var P = 0;
        R = function(B, q, Q) {
          if (B && B.code === "EAGAIN" && P < 10)
            return P++, l.call(e, h, g, _, E, T, R);
          A.apply(this, arguments);
        };
      }
      return l.call(e, h, g, _, E, T, R);
    }
    return Object.setPrototypeOf && Object.setPrototypeOf(f, l), f;
  }(e.read), e.readSync = typeof e.readSync != "function" ? e.readSync : /* @__PURE__ */ function(l) {
    return function(f, h, g, _, E) {
      for (var T = 0; ; )
        try {
          return l.call(e, f, h, g, _, E);
        } catch (A) {
          if (A.code === "EAGAIN" && T < 10) {
            T++;
            continue;
          }
          throw A;
        }
    };
  }(e.readSync);
  function t(l) {
    l.lchmod = function(f, h, g) {
      l.open(
        f,
        pt.O_WRONLY | pt.O_SYMLINK,
        h,
        function(_, E) {
          if (_) {
            g && g(_);
            return;
          }
          l.fchmod(E, h, function(T) {
            l.close(E, function(A) {
              g && g(T || A);
            });
          });
        }
      );
    }, l.lchmodSync = function(f, h) {
      var g = l.openSync(f, pt.O_WRONLY | pt.O_SYMLINK, h), _ = !0, E;
      try {
        E = l.fchmodSync(g, h), _ = !1;
      } finally {
        if (_)
          try {
            l.closeSync(g);
          } catch {
          }
        else
          l.closeSync(g);
      }
      return E;
    };
  }
  function n(l) {
    pt.hasOwnProperty("O_SYMLINK") && l.futimes ? (l.lutimes = function(f, h, g, _) {
      l.open(f, pt.O_SYMLINK, function(E, T) {
        if (E) {
          _ && _(E);
          return;
        }
        l.futimes(T, h, g, function(A) {
          l.close(T, function(R) {
            _ && _(A || R);
          });
        });
      });
    }, l.lutimesSync = function(f, h, g) {
      var _ = l.openSync(f, pt.O_SYMLINK), E, T = !0;
      try {
        E = l.futimesSync(_, h, g), T = !1;
      } finally {
        if (T)
          try {
            l.closeSync(_);
          } catch {
          }
        else
          l.closeSync(_);
      }
      return E;
    }) : l.futimes && (l.lutimes = function(f, h, g, _) {
      _ && process.nextTick(_);
    }, l.lutimesSync = function() {
    });
  }
  function r(l) {
    return l && function(f, h, g) {
      return l.call(e, f, h, function(_) {
        p(_) && (_ = null), g && g.apply(this, arguments);
      });
    };
  }
  function i(l) {
    return l && function(f, h) {
      try {
        return l.call(e, f, h);
      } catch (g) {
        if (!p(g)) throw g;
      }
    };
  }
  function o(l) {
    return l && function(f, h, g, _) {
      return l.call(e, f, h, g, function(E) {
        p(E) && (E = null), _ && _.apply(this, arguments);
      });
    };
  }
  function a(l) {
    return l && function(f, h, g) {
      try {
        return l.call(e, f, h, g);
      } catch (_) {
        if (!p(_)) throw _;
      }
    };
  }
  function s(l) {
    return l && function(f, h, g) {
      typeof h == "function" && (g = h, h = null);
      function _(E, T) {
        T && (T.uid < 0 && (T.uid += 4294967296), T.gid < 0 && (T.gid += 4294967296)), g && g.apply(this, arguments);
      }
      return h ? l.call(e, f, h, _) : l.call(e, f, _);
    };
  }
  function c(l) {
    return l && function(f, h) {
      var g = h ? l.call(e, f, h) : l.call(e, f);
      return g && (g.uid < 0 && (g.uid += 4294967296), g.gid < 0 && (g.gid += 4294967296)), g;
    };
  }
  function p(l) {
    if (!l || l.code === "ENOSYS")
      return !0;
    var f = !process.getuid || process.getuid() !== 0;
    return !!(f && (l.code === "EINVAL" || l.code === "EPERM"));
  }
}
var Ua = Zn.Stream, vd = Td;
function Td(e) {
  return {
    ReadStream: t,
    WriteStream: n
  };
  function t(r, i) {
    if (!(this instanceof t)) return new t(r, i);
    Ua.call(this);
    var o = this;
    this.path = r, this.fd = null, this.readable = !0, this.paused = !1, this.flags = "r", this.mode = 438, this.bufferSize = 64 * 1024, i = i || {};
    for (var a = Object.keys(i), s = 0, c = a.length; s < c; s++) {
      var p = a[s];
      this[p] = i[p];
    }
    if (this.encoding && this.setEncoding(this.encoding), this.start !== void 0) {
      if (typeof this.start != "number")
        throw TypeError("start must be a Number");
      if (this.end === void 0)
        this.end = 1 / 0;
      else if (typeof this.end != "number")
        throw TypeError("end must be a Number");
      if (this.start > this.end)
        throw new Error("start must be <= end");
      this.pos = this.start;
    }
    if (this.fd !== null) {
      process.nextTick(function() {
        o._read();
      });
      return;
    }
    e.open(this.path, this.flags, this.mode, function(l, f) {
      if (l) {
        o.emit("error", l), o.readable = !1;
        return;
      }
      o.fd = f, o.emit("open", f), o._read();
    });
  }
  function n(r, i) {
    if (!(this instanceof n)) return new n(r, i);
    Ua.call(this), this.path = r, this.fd = null, this.writable = !0, this.flags = "w", this.encoding = "binary", this.mode = 438, this.bytesWritten = 0, i = i || {};
    for (var o = Object.keys(i), a = 0, s = o.length; a < s; a++) {
      var c = o[a];
      this[c] = i[c];
    }
    if (this.start !== void 0) {
      if (typeof this.start != "number")
        throw TypeError("start must be a Number");
      if (this.start < 0)
        throw new Error("start must be >= zero");
      this.pos = this.start;
    }
    this.busy = !1, this._queue = [], this.fd === null && (this._open = e.open, this._queue.push([this._open, this.path, this.flags, this.mode, void 0]), this.flush());
  }
}
var Ad = Cd, Sd = Object.getPrototypeOf || function(e) {
  return e.__proto__;
};
function Cd(e) {
  if (e === null || typeof e != "object")
    return e;
  if (e instanceof Object)
    var t = { __proto__: Sd(e) };
  else
    var t = /* @__PURE__ */ Object.create(null);
  return Object.getOwnPropertyNames(e).forEach(function(n) {
    Object.defineProperty(t, n, Object.getOwnPropertyDescriptor(e, n));
  }), t;
}
var ie = Ct, bd = _d, Nd = vd, Rd = Ad, vr = Po, _e, jr;
typeof Symbol == "function" && typeof Symbol.for == "function" ? (_e = Symbol.for("graceful-fs.queue"), jr = Symbol.for("graceful-fs.previous")) : (_e = "___graceful-fs.queue", jr = "___graceful-fs.previous");
function $d() {
}
function jl(e, t) {
  Object.defineProperty(e, _e, {
    get: function() {
      return t;
    }
  });
}
var Bt = $d;
vr.debuglog ? Bt = vr.debuglog("gfs4") : /\bgfs4\b/i.test(process.env.NODE_DEBUG || "") && (Bt = function() {
  var e = vr.format.apply(vr, arguments);
  e = "GFS4: " + e.split(/\n/).join(`
GFS4: `), console.error(e);
});
if (!ie[_e]) {
  var Od = Oe[_e] || [];
  jl(ie, Od), ie.close = function(e) {
    function t(n, r) {
      return e.call(ie, n, function(i) {
        i || ka(), typeof r == "function" && r.apply(this, arguments);
      });
    }
    return Object.defineProperty(t, jr, {
      value: e
    }), t;
  }(ie.close), ie.closeSync = function(e) {
    function t(n) {
      e.apply(ie, arguments), ka();
    }
    return Object.defineProperty(t, jr, {
      value: e
    }), t;
  }(ie.closeSync), /\bgfs4\b/i.test(process.env.NODE_DEBUG || "") && process.on("exit", function() {
    Bt(ie[_e]), Ul.equal(ie[_e].length, 0);
  });
}
Oe[_e] || jl(Oe, ie[_e]);
var Pe = Lo(Rd(ie));
process.env.TEST_GRACEFUL_FS_GLOBAL_PATCH && !ie.__patched && (Pe = Lo(ie), ie.__patched = !0);
function Lo(e) {
  bd(e), e.gracefulify = Lo, e.createReadStream = q, e.createWriteStream = Q;
  var t = e.readFile;
  e.readFile = n;
  function n(U, y, j) {
    return typeof y == "function" && (j = y, y = null), Y(U, y, j);
    function Y(te, $, N, I) {
      return t(te, $, function(b) {
        b && (b.code === "EMFILE" || b.code === "ENFILE") ? Kt([Y, [te, $, N], b, I || Date.now(), Date.now()]) : typeof N == "function" && N.apply(this, arguments);
      });
    }
  }
  var r = e.writeFile;
  e.writeFile = i;
  function i(U, y, j, Y) {
    return typeof j == "function" && (Y = j, j = null), te(U, y, j, Y);
    function te($, N, I, b, D) {
      return r($, N, I, function(O) {
        O && (O.code === "EMFILE" || O.code === "ENFILE") ? Kt([te, [$, N, I, b], O, D || Date.now(), Date.now()]) : typeof b == "function" && b.apply(this, arguments);
      });
    }
  }
  var o = e.appendFile;
  o && (e.appendFile = a);
  function a(U, y, j, Y) {
    return typeof j == "function" && (Y = j, j = null), te(U, y, j, Y);
    function te($, N, I, b, D) {
      return o($, N, I, function(O) {
        O && (O.code === "EMFILE" || O.code === "ENFILE") ? Kt([te, [$, N, I, b], O, D || Date.now(), Date.now()]) : typeof b == "function" && b.apply(this, arguments);
      });
    }
  }
  var s = e.copyFile;
  s && (e.copyFile = c);
  function c(U, y, j, Y) {
    return typeof j == "function" && (Y = j, j = 0), te(U, y, j, Y);
    function te($, N, I, b, D) {
      return s($, N, I, function(O) {
        O && (O.code === "EMFILE" || O.code === "ENFILE") ? Kt([te, [$, N, I, b], O, D || Date.now(), Date.now()]) : typeof b == "function" && b.apply(this, arguments);
      });
    }
  }
  var p = e.readdir;
  e.readdir = f;
  var l = /^v[0-5]\./;
  function f(U, y, j) {
    typeof y == "function" && (j = y, y = null);
    var Y = l.test(process.version) ? function(N, I, b, D) {
      return p(N, te(
        N,
        I,
        b,
        D
      ));
    } : function(N, I, b, D) {
      return p(N, I, te(
        N,
        I,
        b,
        D
      ));
    };
    return Y(U, y, j);
    function te($, N, I, b) {
      return function(D, O) {
        D && (D.code === "EMFILE" || D.code === "ENFILE") ? Kt([
          Y,
          [$, N, I],
          D,
          b || Date.now(),
          Date.now()
        ]) : (O && O.sort && O.sort(), typeof I == "function" && I.call(this, D, O));
      };
    }
  }
  if (process.version.substr(0, 4) === "v0.8") {
    var h = Nd(e);
    A = h.ReadStream, P = h.WriteStream;
  }
  var g = e.ReadStream;
  g && (A.prototype = Object.create(g.prototype), A.prototype.open = R);
  var _ = e.WriteStream;
  _ && (P.prototype = Object.create(_.prototype), P.prototype.open = B), Object.defineProperty(e, "ReadStream", {
    get: function() {
      return A;
    },
    set: function(U) {
      A = U;
    },
    enumerable: !0,
    configurable: !0
  }), Object.defineProperty(e, "WriteStream", {
    get: function() {
      return P;
    },
    set: function(U) {
      P = U;
    },
    enumerable: !0,
    configurable: !0
  });
  var E = A;
  Object.defineProperty(e, "FileReadStream", {
    get: function() {
      return E;
    },
    set: function(U) {
      E = U;
    },
    enumerable: !0,
    configurable: !0
  });
  var T = P;
  Object.defineProperty(e, "FileWriteStream", {
    get: function() {
      return T;
    },
    set: function(U) {
      T = U;
    },
    enumerable: !0,
    configurable: !0
  });
  function A(U, y) {
    return this instanceof A ? (g.apply(this, arguments), this) : A.apply(Object.create(A.prototype), arguments);
  }
  function R() {
    var U = this;
    ae(U.path, U.flags, U.mode, function(y, j) {
      y ? (U.autoClose && U.destroy(), U.emit("error", y)) : (U.fd = j, U.emit("open", j), U.read());
    });
  }
  function P(U, y) {
    return this instanceof P ? (_.apply(this, arguments), this) : P.apply(Object.create(P.prototype), arguments);
  }
  function B() {
    var U = this;
    ae(U.path, U.flags, U.mode, function(y, j) {
      y ? (U.destroy(), U.emit("error", y)) : (U.fd = j, U.emit("open", j));
    });
  }
  function q(U, y) {
    return new e.ReadStream(U, y);
  }
  function Q(U, y) {
    return new e.WriteStream(U, y);
  }
  var Z = e.open;
  e.open = ae;
  function ae(U, y, j, Y) {
    return typeof j == "function" && (Y = j, j = null), te(U, y, j, Y);
    function te($, N, I, b, D) {
      return Z($, N, I, function(O, k) {
        O && (O.code === "EMFILE" || O.code === "ENFILE") ? Kt([te, [$, N, I, b], O, D || Date.now(), Date.now()]) : typeof b == "function" && b.apply(this, arguments);
      });
    }
  }
  return e;
}
function Kt(e) {
  Bt("ENQUEUE", e[0].name, e[1]), ie[_e].push(e), Fo();
}
var Tr;
function ka() {
  for (var e = Date.now(), t = 0; t < ie[_e].length; ++t)
    ie[_e][t].length > 2 && (ie[_e][t][3] = e, ie[_e][t][4] = e);
  Fo();
}
function Fo() {
  if (clearTimeout(Tr), Tr = void 0, ie[_e].length !== 0) {
    var e = ie[_e].shift(), t = e[0], n = e[1], r = e[2], i = e[3], o = e[4];
    if (i === void 0)
      Bt("RETRY", t.name, n), t.apply(null, n);
    else if (Date.now() - i >= 6e4) {
      Bt("TIMEOUT", t.name, n);
      var a = n.pop();
      typeof a == "function" && a.call(null, r);
    } else {
      var s = Date.now() - o, c = Math.max(o - i, 1), p = Math.min(c * 1.2, 100);
      s >= p ? (Bt("RETRY", t.name, n), t.apply(null, n.concat([i]))) : ie[_e].push(e);
    }
    Tr === void 0 && (Tr = setTimeout(Fo, 0));
  }
}
(function(e) {
  const t = De.fromCallback, n = Pe, r = [
    "access",
    "appendFile",
    "chmod",
    "chown",
    "close",
    "copyFile",
    "fchmod",
    "fchown",
    "fdatasync",
    "fstat",
    "fsync",
    "ftruncate",
    "futimes",
    "lchmod",
    "lchown",
    "link",
    "lstat",
    "mkdir",
    "mkdtemp",
    "open",
    "opendir",
    "readdir",
    "readFile",
    "readlink",
    "realpath",
    "rename",
    "rm",
    "rmdir",
    "stat",
    "symlink",
    "truncate",
    "unlink",
    "utimes",
    "writeFile"
  ].filter((i) => typeof n[i] == "function");
  Object.assign(e, n), r.forEach((i) => {
    e[i] = t(n[i]);
  }), e.exists = function(i, o) {
    return typeof o == "function" ? n.exists(i, o) : new Promise((a) => n.exists(i, a));
  }, e.read = function(i, o, a, s, c, p) {
    return typeof p == "function" ? n.read(i, o, a, s, c, p) : new Promise((l, f) => {
      n.read(i, o, a, s, c, (h, g, _) => {
        if (h) return f(h);
        l({ bytesRead: g, buffer: _ });
      });
    });
  }, e.write = function(i, o, ...a) {
    return typeof a[a.length - 1] == "function" ? n.write(i, o, ...a) : new Promise((s, c) => {
      n.write(i, o, ...a, (p, l, f) => {
        if (p) return c(p);
        s({ bytesWritten: l, buffer: f });
      });
    });
  }, typeof n.writev == "function" && (e.writev = function(i, o, ...a) {
    return typeof a[a.length - 1] == "function" ? n.writev(i, o, ...a) : new Promise((s, c) => {
      n.writev(i, o, ...a, (p, l, f) => {
        if (p) return c(p);
        s({ bytesWritten: l, buffers: f });
      });
    });
  }), typeof n.realpath.native == "function" ? e.realpath.native = t(n.realpath.native) : process.emitWarning(
    "fs.realpath.native is not a function. Is fs being monkey-patched?",
    "Warning",
    "fs-extra-WARN0003"
  );
})(Gt);
var xo = {}, Gl = {};
const Id = ee;
Gl.checkPath = function(t) {
  if (process.platform === "win32" && /[<>:"|?*]/.test(t.replace(Id.parse(t).root, ""))) {
    const r = new Error(`Path contains invalid characters: ${t}`);
    throw r.code = "EINVAL", r;
  }
};
const ql = Gt, { checkPath: Wl } = Gl, Vl = (e) => {
  const t = { mode: 511 };
  return typeof e == "number" ? e : { ...t, ...e }.mode;
};
xo.makeDir = async (e, t) => (Wl(e), ql.mkdir(e, {
  mode: Vl(t),
  recursive: !0
}));
xo.makeDirSync = (e, t) => (Wl(e), ql.mkdirSync(e, {
  mode: Vl(t),
  recursive: !0
}));
const Dd = De.fromPromise, { makeDir: Pd, makeDirSync: Li } = xo, Fi = Dd(Pd);
var it = {
  mkdirs: Fi,
  mkdirsSync: Li,
  // alias
  mkdirp: Fi,
  mkdirpSync: Li,
  ensureDir: Fi,
  ensureDirSync: Li
};
const Ld = De.fromPromise, zl = Gt;
function Fd(e) {
  return zl.access(e).then(() => !0).catch(() => !1);
}
var qt = {
  pathExists: Ld(Fd),
  pathExistsSync: zl.existsSync
};
const cn = Pe;
function xd(e, t, n, r) {
  cn.open(e, "r+", (i, o) => {
    if (i) return r(i);
    cn.futimes(o, t, n, (a) => {
      cn.close(o, (s) => {
        r && r(a || s);
      });
    });
  });
}
function Ud(e, t, n) {
  const r = cn.openSync(e, "r+");
  return cn.futimesSync(r, t, n), cn.closeSync(r);
}
var Yl = {
  utimesMillis: xd,
  utimesMillisSync: Ud
};
const fn = Gt, ge = ee, kd = Po;
function Md(e, t, n) {
  const r = n.dereference ? (i) => fn.stat(i, { bigint: !0 }) : (i) => fn.lstat(i, { bigint: !0 });
  return Promise.all([
    r(e),
    r(t).catch((i) => {
      if (i.code === "ENOENT") return null;
      throw i;
    })
  ]).then(([i, o]) => ({ srcStat: i, destStat: o }));
}
function Bd(e, t, n) {
  let r;
  const i = n.dereference ? (a) => fn.statSync(a, { bigint: !0 }) : (a) => fn.lstatSync(a, { bigint: !0 }), o = i(e);
  try {
    r = i(t);
  } catch (a) {
    if (a.code === "ENOENT") return { srcStat: o, destStat: null };
    throw a;
  }
  return { srcStat: o, destStat: r };
}
function Hd(e, t, n, r, i) {
  kd.callbackify(Md)(e, t, r, (o, a) => {
    if (o) return i(o);
    const { srcStat: s, destStat: c } = a;
    if (c) {
      if (tr(s, c)) {
        const p = ge.basename(e), l = ge.basename(t);
        return n === "move" && p !== l && p.toLowerCase() === l.toLowerCase() ? i(null, { srcStat: s, destStat: c, isChangingCase: !0 }) : i(new Error("Source and destination must not be the same."));
      }
      if (s.isDirectory() && !c.isDirectory())
        return i(new Error(`Cannot overwrite non-directory '${t}' with directory '${e}'.`));
      if (!s.isDirectory() && c.isDirectory())
        return i(new Error(`Cannot overwrite directory '${t}' with non-directory '${e}'.`));
    }
    return s.isDirectory() && Uo(e, t) ? i(new Error(ni(e, t, n))) : i(null, { srcStat: s, destStat: c });
  });
}
function jd(e, t, n, r) {
  const { srcStat: i, destStat: o } = Bd(e, t, r);
  if (o) {
    if (tr(i, o)) {
      const a = ge.basename(e), s = ge.basename(t);
      if (n === "move" && a !== s && a.toLowerCase() === s.toLowerCase())
        return { srcStat: i, destStat: o, isChangingCase: !0 };
      throw new Error("Source and destination must not be the same.");
    }
    if (i.isDirectory() && !o.isDirectory())
      throw new Error(`Cannot overwrite non-directory '${t}' with directory '${e}'.`);
    if (!i.isDirectory() && o.isDirectory())
      throw new Error(`Cannot overwrite directory '${t}' with non-directory '${e}'.`);
  }
  if (i.isDirectory() && Uo(e, t))
    throw new Error(ni(e, t, n));
  return { srcStat: i, destStat: o };
}
function Xl(e, t, n, r, i) {
  const o = ge.resolve(ge.dirname(e)), a = ge.resolve(ge.dirname(n));
  if (a === o || a === ge.parse(a).root) return i();
  fn.stat(a, { bigint: !0 }, (s, c) => s ? s.code === "ENOENT" ? i() : i(s) : tr(t, c) ? i(new Error(ni(e, n, r))) : Xl(e, t, a, r, i));
}
function Kl(e, t, n, r) {
  const i = ge.resolve(ge.dirname(e)), o = ge.resolve(ge.dirname(n));
  if (o === i || o === ge.parse(o).root) return;
  let a;
  try {
    a = fn.statSync(o, { bigint: !0 });
  } catch (s) {
    if (s.code === "ENOENT") return;
    throw s;
  }
  if (tr(t, a))
    throw new Error(ni(e, n, r));
  return Kl(e, t, o, r);
}
function tr(e, t) {
  return t.ino && t.dev && t.ino === e.ino && t.dev === e.dev;
}
function Uo(e, t) {
  const n = ge.resolve(e).split(ge.sep).filter((i) => i), r = ge.resolve(t).split(ge.sep).filter((i) => i);
  return n.reduce((i, o, a) => i && r[a] === o, !0);
}
function ni(e, t, n) {
  return `Cannot ${n} '${e}' to a subdirectory of itself, '${t}'.`;
}
var mn = {
  checkPaths: Hd,
  checkPathsSync: jd,
  checkParentPaths: Xl,
  checkParentPathsSync: Kl,
  isSrcSubdir: Uo,
  areIdentical: tr
};
const xe = Pe, kn = ee, Gd = it.mkdirs, qd = qt.pathExists, Wd = Yl.utimesMillis, Mn = mn;
function Vd(e, t, n, r) {
  typeof n == "function" && !r ? (r = n, n = {}) : typeof n == "function" && (n = { filter: n }), r = r || function() {
  }, n = n || {}, n.clobber = "clobber" in n ? !!n.clobber : !0, n.overwrite = "overwrite" in n ? !!n.overwrite : n.clobber, n.preserveTimestamps && process.arch === "ia32" && process.emitWarning(
    `Using the preserveTimestamps option in 32-bit node is not recommended;

	see https://github.com/jprichardson/node-fs-extra/issues/269`,
    "Warning",
    "fs-extra-WARN0001"
  ), Mn.checkPaths(e, t, "copy", n, (i, o) => {
    if (i) return r(i);
    const { srcStat: a, destStat: s } = o;
    Mn.checkParentPaths(e, a, t, "copy", (c) => c ? r(c) : n.filter ? Jl(Ma, s, e, t, n, r) : Ma(s, e, t, n, r));
  });
}
function Ma(e, t, n, r, i) {
  const o = kn.dirname(n);
  qd(o, (a, s) => {
    if (a) return i(a);
    if (s) return Gr(e, t, n, r, i);
    Gd(o, (c) => c ? i(c) : Gr(e, t, n, r, i));
  });
}
function Jl(e, t, n, r, i, o) {
  Promise.resolve(i.filter(n, r)).then((a) => a ? e(t, n, r, i, o) : o(), (a) => o(a));
}
function zd(e, t, n, r, i) {
  return r.filter ? Jl(Gr, e, t, n, r, i) : Gr(e, t, n, r, i);
}
function Gr(e, t, n, r, i) {
  (r.dereference ? xe.stat : xe.lstat)(t, (a, s) => a ? i(a) : s.isDirectory() ? eh(s, e, t, n, r, i) : s.isFile() || s.isCharacterDevice() || s.isBlockDevice() ? Yd(s, e, t, n, r, i) : s.isSymbolicLink() ? rh(e, t, n, r, i) : s.isSocket() ? i(new Error(`Cannot copy a socket file: ${t}`)) : s.isFIFO() ? i(new Error(`Cannot copy a FIFO pipe: ${t}`)) : i(new Error(`Unknown file: ${t}`)));
}
function Yd(e, t, n, r, i, o) {
  return t ? Xd(e, n, r, i, o) : Ql(e, n, r, i, o);
}
function Xd(e, t, n, r, i) {
  if (r.overwrite)
    xe.unlink(n, (o) => o ? i(o) : Ql(e, t, n, r, i));
  else return r.errorOnExist ? i(new Error(`'${n}' already exists`)) : i();
}
function Ql(e, t, n, r, i) {
  xe.copyFile(t, n, (o) => o ? i(o) : r.preserveTimestamps ? Kd(e.mode, t, n, i) : ri(n, e.mode, i));
}
function Kd(e, t, n, r) {
  return Jd(e) ? Qd(n, e, (i) => i ? r(i) : Ba(e, t, n, r)) : Ba(e, t, n, r);
}
function Jd(e) {
  return (e & 128) === 0;
}
function Qd(e, t, n) {
  return ri(e, t | 128, n);
}
function Ba(e, t, n, r) {
  Zd(t, n, (i) => i ? r(i) : ri(n, e, r));
}
function ri(e, t, n) {
  return xe.chmod(e, t, n);
}
function Zd(e, t, n) {
  xe.stat(e, (r, i) => r ? n(r) : Wd(t, i.atime, i.mtime, n));
}
function eh(e, t, n, r, i, o) {
  return t ? Zl(n, r, i, o) : th(e.mode, n, r, i, o);
}
function th(e, t, n, r, i) {
  xe.mkdir(n, (o) => {
    if (o) return i(o);
    Zl(t, n, r, (a) => a ? i(a) : ri(n, e, i));
  });
}
function Zl(e, t, n, r) {
  xe.readdir(e, (i, o) => i ? r(i) : ec(o, e, t, n, r));
}
function ec(e, t, n, r, i) {
  const o = e.pop();
  return o ? nh(e, o, t, n, r, i) : i();
}
function nh(e, t, n, r, i, o) {
  const a = kn.join(n, t), s = kn.join(r, t);
  Mn.checkPaths(a, s, "copy", i, (c, p) => {
    if (c) return o(c);
    const { destStat: l } = p;
    zd(l, a, s, i, (f) => f ? o(f) : ec(e, n, r, i, o));
  });
}
function rh(e, t, n, r, i) {
  xe.readlink(t, (o, a) => {
    if (o) return i(o);
    if (r.dereference && (a = kn.resolve(process.cwd(), a)), e)
      xe.readlink(n, (s, c) => s ? s.code === "EINVAL" || s.code === "UNKNOWN" ? xe.symlink(a, n, i) : i(s) : (r.dereference && (c = kn.resolve(process.cwd(), c)), Mn.isSrcSubdir(a, c) ? i(new Error(`Cannot copy '${a}' to a subdirectory of itself, '${c}'.`)) : e.isDirectory() && Mn.isSrcSubdir(c, a) ? i(new Error(`Cannot overwrite '${c}' with '${a}'.`)) : ih(a, n, i)));
    else
      return xe.symlink(a, n, i);
  });
}
function ih(e, t, n) {
  xe.unlink(t, (r) => r ? n(r) : xe.symlink(e, t, n));
}
var oh = Vd;
const be = Pe, Bn = ee, ah = it.mkdirsSync, sh = Yl.utimesMillisSync, Hn = mn;
function lh(e, t, n) {
  typeof n == "function" && (n = { filter: n }), n = n || {}, n.clobber = "clobber" in n ? !!n.clobber : !0, n.overwrite = "overwrite" in n ? !!n.overwrite : n.clobber, n.preserveTimestamps && process.arch === "ia32" && process.emitWarning(
    `Using the preserveTimestamps option in 32-bit node is not recommended;

	see https://github.com/jprichardson/node-fs-extra/issues/269`,
    "Warning",
    "fs-extra-WARN0002"
  );
  const { srcStat: r, destStat: i } = Hn.checkPathsSync(e, t, "copy", n);
  return Hn.checkParentPathsSync(e, r, t, "copy"), ch(i, e, t, n);
}
function ch(e, t, n, r) {
  if (r.filter && !r.filter(t, n)) return;
  const i = Bn.dirname(n);
  return be.existsSync(i) || ah(i), tc(e, t, n, r);
}
function uh(e, t, n, r) {
  if (!(r.filter && !r.filter(t, n)))
    return tc(e, t, n, r);
}
function tc(e, t, n, r) {
  const o = (r.dereference ? be.statSync : be.lstatSync)(t);
  if (o.isDirectory()) return Eh(o, e, t, n, r);
  if (o.isFile() || o.isCharacterDevice() || o.isBlockDevice()) return fh(o, e, t, n, r);
  if (o.isSymbolicLink()) return wh(e, t, n, r);
  throw o.isSocket() ? new Error(`Cannot copy a socket file: ${t}`) : o.isFIFO() ? new Error(`Cannot copy a FIFO pipe: ${t}`) : new Error(`Unknown file: ${t}`);
}
function fh(e, t, n, r, i) {
  return t ? dh(e, n, r, i) : nc(e, n, r, i);
}
function dh(e, t, n, r) {
  if (r.overwrite)
    return be.unlinkSync(n), nc(e, t, n, r);
  if (r.errorOnExist)
    throw new Error(`'${n}' already exists`);
}
function nc(e, t, n, r) {
  return be.copyFileSync(t, n), r.preserveTimestamps && hh(e.mode, t, n), ko(n, e.mode);
}
function hh(e, t, n) {
  return ph(e) && mh(n, e), gh(t, n);
}
function ph(e) {
  return (e & 128) === 0;
}
function mh(e, t) {
  return ko(e, t | 128);
}
function ko(e, t) {
  return be.chmodSync(e, t);
}
function gh(e, t) {
  const n = be.statSync(e);
  return sh(t, n.atime, n.mtime);
}
function Eh(e, t, n, r, i) {
  return t ? rc(n, r, i) : yh(e.mode, n, r, i);
}
function yh(e, t, n, r) {
  return be.mkdirSync(n), rc(t, n, r), ko(n, e);
}
function rc(e, t, n) {
  be.readdirSync(e).forEach((r) => _h(r, e, t, n));
}
function _h(e, t, n, r) {
  const i = Bn.join(t, e), o = Bn.join(n, e), { destStat: a } = Hn.checkPathsSync(i, o, "copy", r);
  return uh(a, i, o, r);
}
function wh(e, t, n, r) {
  let i = be.readlinkSync(t);
  if (r.dereference && (i = Bn.resolve(process.cwd(), i)), e) {
    let o;
    try {
      o = be.readlinkSync(n);
    } catch (a) {
      if (a.code === "EINVAL" || a.code === "UNKNOWN") return be.symlinkSync(i, n);
      throw a;
    }
    if (r.dereference && (o = Bn.resolve(process.cwd(), o)), Hn.isSrcSubdir(i, o))
      throw new Error(`Cannot copy '${i}' to a subdirectory of itself, '${o}'.`);
    if (be.statSync(n).isDirectory() && Hn.isSrcSubdir(o, i))
      throw new Error(`Cannot overwrite '${o}' with '${i}'.`);
    return vh(i, n);
  } else
    return be.symlinkSync(i, n);
}
function vh(e, t) {
  return be.unlinkSync(t), be.symlinkSync(e, t);
}
var Th = lh;
const Ah = De.fromCallback;
var Mo = {
  copy: Ah(oh),
  copySync: Th
};
const Ha = Pe, ic = ee, K = Ul, jn = process.platform === "win32";
function oc(e) {
  [
    "unlink",
    "chmod",
    "stat",
    "lstat",
    "rmdir",
    "readdir"
  ].forEach((n) => {
    e[n] = e[n] || Ha[n], n = n + "Sync", e[n] = e[n] || Ha[n];
  }), e.maxBusyTries = e.maxBusyTries || 3;
}
function Bo(e, t, n) {
  let r = 0;
  typeof t == "function" && (n = t, t = {}), K(e, "rimraf: missing path"), K.strictEqual(typeof e, "string", "rimraf: path should be a string"), K.strictEqual(typeof n, "function", "rimraf: callback function required"), K(t, "rimraf: invalid options argument provided"), K.strictEqual(typeof t, "object", "rimraf: options should be object"), oc(t), ja(e, t, function i(o) {
    if (o) {
      if ((o.code === "EBUSY" || o.code === "ENOTEMPTY" || o.code === "EPERM") && r < t.maxBusyTries) {
        r++;
        const a = r * 100;
        return setTimeout(() => ja(e, t, i), a);
      }
      o.code === "ENOENT" && (o = null);
    }
    n(o);
  });
}
function ja(e, t, n) {
  K(e), K(t), K(typeof n == "function"), t.lstat(e, (r, i) => {
    if (r && r.code === "ENOENT")
      return n(null);
    if (r && r.code === "EPERM" && jn)
      return Ga(e, t, r, n);
    if (i && i.isDirectory())
      return Mr(e, t, r, n);
    t.unlink(e, (o) => {
      if (o) {
        if (o.code === "ENOENT")
          return n(null);
        if (o.code === "EPERM")
          return jn ? Ga(e, t, o, n) : Mr(e, t, o, n);
        if (o.code === "EISDIR")
          return Mr(e, t, o, n);
      }
      return n(o);
    });
  });
}
function Ga(e, t, n, r) {
  K(e), K(t), K(typeof r == "function"), t.chmod(e, 438, (i) => {
    i ? r(i.code === "ENOENT" ? null : n) : t.stat(e, (o, a) => {
      o ? r(o.code === "ENOENT" ? null : n) : a.isDirectory() ? Mr(e, t, n, r) : t.unlink(e, r);
    });
  });
}
function qa(e, t, n) {
  let r;
  K(e), K(t);
  try {
    t.chmodSync(e, 438);
  } catch (i) {
    if (i.code === "ENOENT")
      return;
    throw n;
  }
  try {
    r = t.statSync(e);
  } catch (i) {
    if (i.code === "ENOENT")
      return;
    throw n;
  }
  r.isDirectory() ? Br(e, t, n) : t.unlinkSync(e);
}
function Mr(e, t, n, r) {
  K(e), K(t), K(typeof r == "function"), t.rmdir(e, (i) => {
    i && (i.code === "ENOTEMPTY" || i.code === "EEXIST" || i.code === "EPERM") ? Sh(e, t, r) : i && i.code === "ENOTDIR" ? r(n) : r(i);
  });
}
function Sh(e, t, n) {
  K(e), K(t), K(typeof n == "function"), t.readdir(e, (r, i) => {
    if (r) return n(r);
    let o = i.length, a;
    if (o === 0) return t.rmdir(e, n);
    i.forEach((s) => {
      Bo(ic.join(e, s), t, (c) => {
        if (!a) {
          if (c) return n(a = c);
          --o === 0 && t.rmdir(e, n);
        }
      });
    });
  });
}
function ac(e, t) {
  let n;
  t = t || {}, oc(t), K(e, "rimraf: missing path"), K.strictEqual(typeof e, "string", "rimraf: path should be a string"), K(t, "rimraf: missing options"), K.strictEqual(typeof t, "object", "rimraf: options should be object");
  try {
    n = t.lstatSync(e);
  } catch (r) {
    if (r.code === "ENOENT")
      return;
    r.code === "EPERM" && jn && qa(e, t, r);
  }
  try {
    n && n.isDirectory() ? Br(e, t, null) : t.unlinkSync(e);
  } catch (r) {
    if (r.code === "ENOENT")
      return;
    if (r.code === "EPERM")
      return jn ? qa(e, t, r) : Br(e, t, r);
    if (r.code !== "EISDIR")
      throw r;
    Br(e, t, r);
  }
}
function Br(e, t, n) {
  K(e), K(t);
  try {
    t.rmdirSync(e);
  } catch (r) {
    if (r.code === "ENOTDIR")
      throw n;
    if (r.code === "ENOTEMPTY" || r.code === "EEXIST" || r.code === "EPERM")
      Ch(e, t);
    else if (r.code !== "ENOENT")
      throw r;
  }
}
function Ch(e, t) {
  if (K(e), K(t), t.readdirSync(e).forEach((n) => ac(ic.join(e, n), t)), jn) {
    const n = Date.now();
    do
      try {
        return t.rmdirSync(e, t);
      } catch {
      }
    while (Date.now() - n < 500);
  } else
    return t.rmdirSync(e, t);
}
var bh = Bo;
Bo.sync = ac;
const qr = Pe, Nh = De.fromCallback, sc = bh;
function Rh(e, t) {
  if (qr.rm) return qr.rm(e, { recursive: !0, force: !0 }, t);
  sc(e, t);
}
function $h(e) {
  if (qr.rmSync) return qr.rmSync(e, { recursive: !0, force: !0 });
  sc.sync(e);
}
var ii = {
  remove: Nh(Rh),
  removeSync: $h
};
const Oh = De.fromPromise, lc = Gt, cc = ee, uc = it, fc = ii, Wa = Oh(async function(t) {
  let n;
  try {
    n = await lc.readdir(t);
  } catch {
    return uc.mkdirs(t);
  }
  return Promise.all(n.map((r) => fc.remove(cc.join(t, r))));
});
function Va(e) {
  let t;
  try {
    t = lc.readdirSync(e);
  } catch {
    return uc.mkdirsSync(e);
  }
  t.forEach((n) => {
    n = cc.join(e, n), fc.removeSync(n);
  });
}
var Ih = {
  emptyDirSync: Va,
  emptydirSync: Va,
  emptyDir: Wa,
  emptydir: Wa
};
const Dh = De.fromCallback, dc = ee, yt = Pe, hc = it;
function Ph(e, t) {
  function n() {
    yt.writeFile(e, "", (r) => {
      if (r) return t(r);
      t();
    });
  }
  yt.stat(e, (r, i) => {
    if (!r && i.isFile()) return t();
    const o = dc.dirname(e);
    yt.stat(o, (a, s) => {
      if (a)
        return a.code === "ENOENT" ? hc.mkdirs(o, (c) => {
          if (c) return t(c);
          n();
        }) : t(a);
      s.isDirectory() ? n() : yt.readdir(o, (c) => {
        if (c) return t(c);
      });
    });
  });
}
function Lh(e) {
  let t;
  try {
    t = yt.statSync(e);
  } catch {
  }
  if (t && t.isFile()) return;
  const n = dc.dirname(e);
  try {
    yt.statSync(n).isDirectory() || yt.readdirSync(n);
  } catch (r) {
    if (r && r.code === "ENOENT") hc.mkdirsSync(n);
    else throw r;
  }
  yt.writeFileSync(e, "");
}
var Fh = {
  createFile: Dh(Ph),
  createFileSync: Lh
};
const xh = De.fromCallback, pc = ee, Et = Pe, mc = it, Uh = qt.pathExists, { areIdentical: gc } = mn;
function kh(e, t, n) {
  function r(i, o) {
    Et.link(i, o, (a) => {
      if (a) return n(a);
      n(null);
    });
  }
  Et.lstat(t, (i, o) => {
    Et.lstat(e, (a, s) => {
      if (a)
        return a.message = a.message.replace("lstat", "ensureLink"), n(a);
      if (o && gc(s, o)) return n(null);
      const c = pc.dirname(t);
      Uh(c, (p, l) => {
        if (p) return n(p);
        if (l) return r(e, t);
        mc.mkdirs(c, (f) => {
          if (f) return n(f);
          r(e, t);
        });
      });
    });
  });
}
function Mh(e, t) {
  let n;
  try {
    n = Et.lstatSync(t);
  } catch {
  }
  try {
    const o = Et.lstatSync(e);
    if (n && gc(o, n)) return;
  } catch (o) {
    throw o.message = o.message.replace("lstat", "ensureLink"), o;
  }
  const r = pc.dirname(t);
  return Et.existsSync(r) || mc.mkdirsSync(r), Et.linkSync(e, t);
}
var Bh = {
  createLink: xh(kh),
  createLinkSync: Mh
};
const _t = ee, Ln = Pe, Hh = qt.pathExists;
function jh(e, t, n) {
  if (_t.isAbsolute(e))
    return Ln.lstat(e, (r) => r ? (r.message = r.message.replace("lstat", "ensureSymlink"), n(r)) : n(null, {
      toCwd: e,
      toDst: e
    }));
  {
    const r = _t.dirname(t), i = _t.join(r, e);
    return Hh(i, (o, a) => o ? n(o) : a ? n(null, {
      toCwd: i,
      toDst: e
    }) : Ln.lstat(e, (s) => s ? (s.message = s.message.replace("lstat", "ensureSymlink"), n(s)) : n(null, {
      toCwd: e,
      toDst: _t.relative(r, e)
    })));
  }
}
function Gh(e, t) {
  let n;
  if (_t.isAbsolute(e)) {
    if (n = Ln.existsSync(e), !n) throw new Error("absolute srcpath does not exist");
    return {
      toCwd: e,
      toDst: e
    };
  } else {
    const r = _t.dirname(t), i = _t.join(r, e);
    if (n = Ln.existsSync(i), n)
      return {
        toCwd: i,
        toDst: e
      };
    if (n = Ln.existsSync(e), !n) throw new Error("relative srcpath does not exist");
    return {
      toCwd: e,
      toDst: _t.relative(r, e)
    };
  }
}
var qh = {
  symlinkPaths: jh,
  symlinkPathsSync: Gh
};
const Ec = Pe;
function Wh(e, t, n) {
  if (n = typeof t == "function" ? t : n, t = typeof t == "function" ? !1 : t, t) return n(null, t);
  Ec.lstat(e, (r, i) => {
    if (r) return n(null, "file");
    t = i && i.isDirectory() ? "dir" : "file", n(null, t);
  });
}
function Vh(e, t) {
  let n;
  if (t) return t;
  try {
    n = Ec.lstatSync(e);
  } catch {
    return "file";
  }
  return n && n.isDirectory() ? "dir" : "file";
}
var zh = {
  symlinkType: Wh,
  symlinkTypeSync: Vh
};
const Yh = De.fromCallback, yc = ee, Ye = Gt, _c = it, Xh = _c.mkdirs, Kh = _c.mkdirsSync, wc = qh, Jh = wc.symlinkPaths, Qh = wc.symlinkPathsSync, vc = zh, Zh = vc.symlinkType, ep = vc.symlinkTypeSync, tp = qt.pathExists, { areIdentical: Tc } = mn;
function np(e, t, n, r) {
  r = typeof n == "function" ? n : r, n = typeof n == "function" ? !1 : n, Ye.lstat(t, (i, o) => {
    !i && o.isSymbolicLink() ? Promise.all([
      Ye.stat(e),
      Ye.stat(t)
    ]).then(([a, s]) => {
      if (Tc(a, s)) return r(null);
      za(e, t, n, r);
    }) : za(e, t, n, r);
  });
}
function za(e, t, n, r) {
  Jh(e, t, (i, o) => {
    if (i) return r(i);
    e = o.toDst, Zh(o.toCwd, n, (a, s) => {
      if (a) return r(a);
      const c = yc.dirname(t);
      tp(c, (p, l) => {
        if (p) return r(p);
        if (l) return Ye.symlink(e, t, s, r);
        Xh(c, (f) => {
          if (f) return r(f);
          Ye.symlink(e, t, s, r);
        });
      });
    });
  });
}
function rp(e, t, n) {
  let r;
  try {
    r = Ye.lstatSync(t);
  } catch {
  }
  if (r && r.isSymbolicLink()) {
    const s = Ye.statSync(e), c = Ye.statSync(t);
    if (Tc(s, c)) return;
  }
  const i = Qh(e, t);
  e = i.toDst, n = ep(i.toCwd, n);
  const o = yc.dirname(t);
  return Ye.existsSync(o) || Kh(o), Ye.symlinkSync(e, t, n);
}
var ip = {
  createSymlink: Yh(np),
  createSymlinkSync: rp
};
const { createFile: Ya, createFileSync: Xa } = Fh, { createLink: Ka, createLinkSync: Ja } = Bh, { createSymlink: Qa, createSymlinkSync: Za } = ip;
var op = {
  // file
  createFile: Ya,
  createFileSync: Xa,
  ensureFile: Ya,
  ensureFileSync: Xa,
  // link
  createLink: Ka,
  createLinkSync: Ja,
  ensureLink: Ka,
  ensureLinkSync: Ja,
  // symlink
  createSymlink: Qa,
  createSymlinkSync: Za,
  ensureSymlink: Qa,
  ensureSymlinkSync: Za
};
function ap(e, { EOL: t = `
`, finalEOL: n = !0, replacer: r = null, spaces: i } = {}) {
  const o = n ? t : "", a = JSON.stringify(e, r, i);
  if (a === void 0)
    throw new TypeError(`Converting ${typeof e} value to JSON is not supported`);
  return a.replace(/\n/g, t) + o;
}
function sp(e) {
  return Buffer.isBuffer(e) && (e = e.toString("utf8")), e.replace(/^\uFEFF/, "");
}
var Ho = { stringify: ap, stripBom: sp };
let dn;
try {
  dn = Pe;
} catch {
  dn = Ct;
}
const oi = De, { stringify: Ac, stripBom: Sc } = Ho;
async function lp(e, t = {}) {
  typeof t == "string" && (t = { encoding: t });
  const n = t.fs || dn, r = "throws" in t ? t.throws : !0;
  let i = await oi.fromCallback(n.readFile)(e, t);
  i = Sc(i);
  let o;
  try {
    o = JSON.parse(i, t ? t.reviver : null);
  } catch (a) {
    if (r)
      throw a.message = `${e}: ${a.message}`, a;
    return null;
  }
  return o;
}
const cp = oi.fromPromise(lp);
function up(e, t = {}) {
  typeof t == "string" && (t = { encoding: t });
  const n = t.fs || dn, r = "throws" in t ? t.throws : !0;
  try {
    let i = n.readFileSync(e, t);
    return i = Sc(i), JSON.parse(i, t.reviver);
  } catch (i) {
    if (r)
      throw i.message = `${e}: ${i.message}`, i;
    return null;
  }
}
async function fp(e, t, n = {}) {
  const r = n.fs || dn, i = Ac(t, n);
  await oi.fromCallback(r.writeFile)(e, i, n);
}
const dp = oi.fromPromise(fp);
function hp(e, t, n = {}) {
  const r = n.fs || dn, i = Ac(t, n);
  return r.writeFileSync(e, i, n);
}
var pp = {
  readFile: cp,
  readFileSync: up,
  writeFile: dp,
  writeFileSync: hp
};
const Ar = pp;
var mp = {
  // jsonfile exports
  readJson: Ar.readFile,
  readJsonSync: Ar.readFileSync,
  writeJson: Ar.writeFile,
  writeJsonSync: Ar.writeFileSync
};
const gp = De.fromCallback, Fn = Pe, Cc = ee, bc = it, Ep = qt.pathExists;
function yp(e, t, n, r) {
  typeof n == "function" && (r = n, n = "utf8");
  const i = Cc.dirname(e);
  Ep(i, (o, a) => {
    if (o) return r(o);
    if (a) return Fn.writeFile(e, t, n, r);
    bc.mkdirs(i, (s) => {
      if (s) return r(s);
      Fn.writeFile(e, t, n, r);
    });
  });
}
function _p(e, ...t) {
  const n = Cc.dirname(e);
  if (Fn.existsSync(n))
    return Fn.writeFileSync(e, ...t);
  bc.mkdirsSync(n), Fn.writeFileSync(e, ...t);
}
var jo = {
  outputFile: gp(yp),
  outputFileSync: _p
};
const { stringify: wp } = Ho, { outputFile: vp } = jo;
async function Tp(e, t, n = {}) {
  const r = wp(t, n);
  await vp(e, r, n);
}
var Ap = Tp;
const { stringify: Sp } = Ho, { outputFileSync: Cp } = jo;
function bp(e, t, n) {
  const r = Sp(t, n);
  Cp(e, r, n);
}
var Np = bp;
const Rp = De.fromPromise, Ie = mp;
Ie.outputJson = Rp(Ap);
Ie.outputJsonSync = Np;
Ie.outputJSON = Ie.outputJson;
Ie.outputJSONSync = Ie.outputJsonSync;
Ie.writeJSON = Ie.writeJson;
Ie.writeJSONSync = Ie.writeJsonSync;
Ie.readJSON = Ie.readJson;
Ie.readJSONSync = Ie.readJsonSync;
var $p = Ie;
const Op = Pe, Eo = ee, Ip = Mo.copy, Nc = ii.remove, Dp = it.mkdirp, Pp = qt.pathExists, es = mn;
function Lp(e, t, n, r) {
  typeof n == "function" && (r = n, n = {}), n = n || {};
  const i = n.overwrite || n.clobber || !1;
  es.checkPaths(e, t, "move", n, (o, a) => {
    if (o) return r(o);
    const { srcStat: s, isChangingCase: c = !1 } = a;
    es.checkParentPaths(e, s, t, "move", (p) => {
      if (p) return r(p);
      if (Fp(t)) return ts(e, t, i, c, r);
      Dp(Eo.dirname(t), (l) => l ? r(l) : ts(e, t, i, c, r));
    });
  });
}
function Fp(e) {
  const t = Eo.dirname(e);
  return Eo.parse(t).root === t;
}
function ts(e, t, n, r, i) {
  if (r) return xi(e, t, n, i);
  if (n)
    return Nc(t, (o) => o ? i(o) : xi(e, t, n, i));
  Pp(t, (o, a) => o ? i(o) : a ? i(new Error("dest already exists.")) : xi(e, t, n, i));
}
function xi(e, t, n, r) {
  Op.rename(e, t, (i) => i ? i.code !== "EXDEV" ? r(i) : xp(e, t, n, r) : r());
}
function xp(e, t, n, r) {
  Ip(e, t, {
    overwrite: n,
    errorOnExist: !0
  }, (o) => o ? r(o) : Nc(e, r));
}
var Up = Lp;
const Rc = Pe, yo = ee, kp = Mo.copySync, $c = ii.removeSync, Mp = it.mkdirpSync, ns = mn;
function Bp(e, t, n) {
  n = n || {};
  const r = n.overwrite || n.clobber || !1, { srcStat: i, isChangingCase: o = !1 } = ns.checkPathsSync(e, t, "move", n);
  return ns.checkParentPathsSync(e, i, t, "move"), Hp(t) || Mp(yo.dirname(t)), jp(e, t, r, o);
}
function Hp(e) {
  const t = yo.dirname(e);
  return yo.parse(t).root === t;
}
function jp(e, t, n, r) {
  if (r) return Ui(e, t, n);
  if (n)
    return $c(t), Ui(e, t, n);
  if (Rc.existsSync(t)) throw new Error("dest already exists.");
  return Ui(e, t, n);
}
function Ui(e, t, n) {
  try {
    Rc.renameSync(e, t);
  } catch (r) {
    if (r.code !== "EXDEV") throw r;
    return Gp(e, t, n);
  }
}
function Gp(e, t, n) {
  return kp(e, t, {
    overwrite: n,
    errorOnExist: !0
  }), $c(e);
}
var qp = Bp;
const Wp = De.fromCallback;
var Vp = {
  move: Wp(Up),
  moveSync: qp
}, Nt = {
  // Export promiseified graceful-fs:
  ...Gt,
  // Export extra methods:
  ...Mo,
  ...Ih,
  ...op,
  ...$p,
  ...it,
  ...Vp,
  ...jo,
  ...qt,
  ...ii
}, Wt = {}, vt = {}, he = {}, Tt = {};
Object.defineProperty(Tt, "__esModule", { value: !0 });
Tt.CancellationError = Tt.CancellationToken = void 0;
const zp = kl;
class Yp extends zp.EventEmitter {
  get cancelled() {
    return this._cancelled || this._parent != null && this._parent.cancelled;
  }
  set parent(t) {
    this.removeParentCancelHandler(), this._parent = t, this.parentCancelHandler = () => this.cancel(), this._parent.onCancel(this.parentCancelHandler);
  }
  // babel cannot compile ... correctly for super calls
  constructor(t) {
    super(), this.parentCancelHandler = null, this._parent = null, this._cancelled = !1, t != null && (this.parent = t);
  }
  cancel() {
    this._cancelled = !0, this.emit("cancel");
  }
  onCancel(t) {
    this.cancelled ? t() : this.once("cancel", t);
  }
  createPromise(t) {
    if (this.cancelled)
      return Promise.reject(new _o());
    const n = () => {
      if (r != null)
        try {
          this.removeListener("cancel", r), r = null;
        } catch {
        }
    };
    let r = null;
    return new Promise((i, o) => {
      let a = null;
      if (r = () => {
        try {
          a != null && (a(), a = null);
        } finally {
          o(new _o());
        }
      }, this.cancelled) {
        r();
        return;
      }
      this.onCancel(r), t(i, o, (s) => {
        a = s;
      });
    }).then((i) => (n(), i)).catch((i) => {
      throw n(), i;
    });
  }
  removeParentCancelHandler() {
    const t = this._parent;
    t != null && this.parentCancelHandler != null && (t.removeListener("cancel", this.parentCancelHandler), this.parentCancelHandler = null);
  }
  dispose() {
    try {
      this.removeParentCancelHandler();
    } finally {
      this.removeAllListeners(), this._parent = null;
    }
  }
}
Tt.CancellationToken = Yp;
class _o extends Error {
  constructor() {
    super("cancelled");
  }
}
Tt.CancellationError = _o;
var gn = {};
Object.defineProperty(gn, "__esModule", { value: !0 });
gn.newError = Xp;
function Xp(e, t) {
  const n = new Error(e);
  return n.code = t, n;
}
var de = {}, wo = { exports: {} }, Sr = { exports: {} }, ki, rs;
function Kp() {
  if (rs) return ki;
  rs = 1;
  var e = 1e3, t = e * 60, n = t * 60, r = n * 24, i = r * 7, o = r * 365.25;
  ki = function(l, f) {
    f = f || {};
    var h = typeof l;
    if (h === "string" && l.length > 0)
      return a(l);
    if (h === "number" && isFinite(l))
      return f.long ? c(l) : s(l);
    throw new Error(
      "val is not a non-empty string or a valid number. val=" + JSON.stringify(l)
    );
  };
  function a(l) {
    if (l = String(l), !(l.length > 100)) {
      var f = /^(-?(?:\d+)?\.?\d+) *(milliseconds?|msecs?|ms|seconds?|secs?|s|minutes?|mins?|m|hours?|hrs?|h|days?|d|weeks?|w|years?|yrs?|y)?$/i.exec(
        l
      );
      if (f) {
        var h = parseFloat(f[1]), g = (f[2] || "ms").toLowerCase();
        switch (g) {
          case "years":
          case "year":
          case "yrs":
          case "yr":
          case "y":
            return h * o;
          case "weeks":
          case "week":
          case "w":
            return h * i;
          case "days":
          case "day":
          case "d":
            return h * r;
          case "hours":
          case "hour":
          case "hrs":
          case "hr":
          case "h":
            return h * n;
          case "minutes":
          case "minute":
          case "mins":
          case "min":
          case "m":
            return h * t;
          case "seconds":
          case "second":
          case "secs":
          case "sec":
          case "s":
            return h * e;
          case "milliseconds":
          case "millisecond":
          case "msecs":
          case "msec":
          case "ms":
            return h;
          default:
            return;
        }
      }
    }
  }
  function s(l) {
    var f = Math.abs(l);
    return f >= r ? Math.round(l / r) + "d" : f >= n ? Math.round(l / n) + "h" : f >= t ? Math.round(l / t) + "m" : f >= e ? Math.round(l / e) + "s" : l + "ms";
  }
  function c(l) {
    var f = Math.abs(l);
    return f >= r ? p(l, f, r, "day") : f >= n ? p(l, f, n, "hour") : f >= t ? p(l, f, t, "minute") : f >= e ? p(l, f, e, "second") : l + " ms";
  }
  function p(l, f, h, g) {
    var _ = f >= h * 1.5;
    return Math.round(l / h) + " " + g + (_ ? "s" : "");
  }
  return ki;
}
var Mi, is;
function Oc() {
  if (is) return Mi;
  is = 1;
  function e(t) {
    r.debug = r, r.default = r, r.coerce = p, r.disable = s, r.enable = o, r.enabled = c, r.humanize = Kp(), r.destroy = l, Object.keys(t).forEach((f) => {
      r[f] = t[f];
    }), r.names = [], r.skips = [], r.formatters = {};
    function n(f) {
      let h = 0;
      for (let g = 0; g < f.length; g++)
        h = (h << 5) - h + f.charCodeAt(g), h |= 0;
      return r.colors[Math.abs(h) % r.colors.length];
    }
    r.selectColor = n;
    function r(f) {
      let h, g = null, _, E;
      function T(...A) {
        if (!T.enabled)
          return;
        const R = T, P = Number(/* @__PURE__ */ new Date()), B = P - (h || P);
        R.diff = B, R.prev = h, R.curr = P, h = P, A[0] = r.coerce(A[0]), typeof A[0] != "string" && A.unshift("%O");
        let q = 0;
        A[0] = A[0].replace(/%([a-zA-Z%])/g, (Z, ae) => {
          if (Z === "%%")
            return "%";
          q++;
          const U = r.formatters[ae];
          if (typeof U == "function") {
            const y = A[q];
            Z = U.call(R, y), A.splice(q, 1), q--;
          }
          return Z;
        }), r.formatArgs.call(R, A), (R.log || r.log).apply(R, A);
      }
      return T.namespace = f, T.useColors = r.useColors(), T.color = r.selectColor(f), T.extend = i, T.destroy = r.destroy, Object.defineProperty(T, "enabled", {
        enumerable: !0,
        configurable: !1,
        get: () => g !== null ? g : (_ !== r.namespaces && (_ = r.namespaces, E = r.enabled(f)), E),
        set: (A) => {
          g = A;
        }
      }), typeof r.init == "function" && r.init(T), T;
    }
    function i(f, h) {
      const g = r(this.namespace + (typeof h > "u" ? ":" : h) + f);
      return g.log = this.log, g;
    }
    function o(f) {
      r.save(f), r.namespaces = f, r.names = [], r.skips = [];
      const h = (typeof f == "string" ? f : "").trim().replace(/\s+/g, ",").split(",").filter(Boolean);
      for (const g of h)
        g[0] === "-" ? r.skips.push(g.slice(1)) : r.names.push(g);
    }
    function a(f, h) {
      let g = 0, _ = 0, E = -1, T = 0;
      for (; g < f.length; )
        if (_ < h.length && (h[_] === f[g] || h[_] === "*"))
          h[_] === "*" ? (E = _, T = g, _++) : (g++, _++);
        else if (E !== -1)
          _ = E + 1, T++, g = T;
        else
          return !1;
      for (; _ < h.length && h[_] === "*"; )
        _++;
      return _ === h.length;
    }
    function s() {
      const f = [
        ...r.names,
        ...r.skips.map((h) => "-" + h)
      ].join(",");
      return r.enable(""), f;
    }
    function c(f) {
      for (const h of r.skips)
        if (a(f, h))
          return !1;
      for (const h of r.names)
        if (a(f, h))
          return !0;
      return !1;
    }
    function p(f) {
      return f instanceof Error ? f.stack || f.message : f;
    }
    function l() {
      console.warn("Instance method `debug.destroy()` is deprecated and no longer does anything. It will be removed in the next major version of `debug`.");
    }
    return r.enable(r.load()), r;
  }
  return Mi = e, Mi;
}
var os;
function Jp() {
  return os || (os = 1, function(e, t) {
    t.formatArgs = r, t.save = i, t.load = o, t.useColors = n, t.storage = a(), t.destroy = /* @__PURE__ */ (() => {
      let c = !1;
      return () => {
        c || (c = !0, console.warn("Instance method `debug.destroy()` is deprecated and no longer does anything. It will be removed in the next major version of `debug`."));
      };
    })(), t.colors = [
      "#0000CC",
      "#0000FF",
      "#0033CC",
      "#0033FF",
      "#0066CC",
      "#0066FF",
      "#0099CC",
      "#0099FF",
      "#00CC00",
      "#00CC33",
      "#00CC66",
      "#00CC99",
      "#00CCCC",
      "#00CCFF",
      "#3300CC",
      "#3300FF",
      "#3333CC",
      "#3333FF",
      "#3366CC",
      "#3366FF",
      "#3399CC",
      "#3399FF",
      "#33CC00",
      "#33CC33",
      "#33CC66",
      "#33CC99",
      "#33CCCC",
      "#33CCFF",
      "#6600CC",
      "#6600FF",
      "#6633CC",
      "#6633FF",
      "#66CC00",
      "#66CC33",
      "#9900CC",
      "#9900FF",
      "#9933CC",
      "#9933FF",
      "#99CC00",
      "#99CC33",
      "#CC0000",
      "#CC0033",
      "#CC0066",
      "#CC0099",
      "#CC00CC",
      "#CC00FF",
      "#CC3300",
      "#CC3333",
      "#CC3366",
      "#CC3399",
      "#CC33CC",
      "#CC33FF",
      "#CC6600",
      "#CC6633",
      "#CC9900",
      "#CC9933",
      "#CCCC00",
      "#CCCC33",
      "#FF0000",
      "#FF0033",
      "#FF0066",
      "#FF0099",
      "#FF00CC",
      "#FF00FF",
      "#FF3300",
      "#FF3333",
      "#FF3366",
      "#FF3399",
      "#FF33CC",
      "#FF33FF",
      "#FF6600",
      "#FF6633",
      "#FF9900",
      "#FF9933",
      "#FFCC00",
      "#FFCC33"
    ];
    function n() {
      if (typeof window < "u" && window.process && (window.process.type === "renderer" || window.process.__nwjs))
        return !0;
      if (typeof navigator < "u" && navigator.userAgent && navigator.userAgent.toLowerCase().match(/(edge|trident)\/(\d+)/))
        return !1;
      let c;
      return typeof document < "u" && document.documentElement && document.documentElement.style && document.documentElement.style.WebkitAppearance || // Is firebug? http://stackoverflow.com/a/398120/376773
      typeof window < "u" && window.console && (window.console.firebug || window.console.exception && window.console.table) || // Is firefox >= v31?
      // https://developer.mozilla.org/en-US/docs/Tools/Web_Console#Styling_messages
      typeof navigator < "u" && navigator.userAgent && (c = navigator.userAgent.toLowerCase().match(/firefox\/(\d+)/)) && parseInt(c[1], 10) >= 31 || // Double check webkit in userAgent just in case we are in a worker
      typeof navigator < "u" && navigator.userAgent && navigator.userAgent.toLowerCase().match(/applewebkit\/(\d+)/);
    }
    function r(c) {
      if (c[0] = (this.useColors ? "%c" : "") + this.namespace + (this.useColors ? " %c" : " ") + c[0] + (this.useColors ? "%c " : " ") + "+" + e.exports.humanize(this.diff), !this.useColors)
        return;
      const p = "color: " + this.color;
      c.splice(1, 0, p, "color: inherit");
      let l = 0, f = 0;
      c[0].replace(/%[a-zA-Z%]/g, (h) => {
        h !== "%%" && (l++, h === "%c" && (f = l));
      }), c.splice(f, 0, p);
    }
    t.log = console.debug || console.log || (() => {
    });
    function i(c) {
      try {
        c ? t.storage.setItem("debug", c) : t.storage.removeItem("debug");
      } catch {
      }
    }
    function o() {
      let c;
      try {
        c = t.storage.getItem("debug") || t.storage.getItem("DEBUG");
      } catch {
      }
      return !c && typeof process < "u" && "env" in process && (c = process.env.DEBUG), c;
    }
    function a() {
      try {
        return localStorage;
      } catch {
      }
    }
    e.exports = Oc()(t);
    const { formatters: s } = e.exports;
    s.j = function(c) {
      try {
        return JSON.stringify(c);
      } catch (p) {
        return "[UnexpectedJSONParseError]: " + p.message;
      }
    };
  }(Sr, Sr.exports)), Sr.exports;
}
var Cr = { exports: {} }, Bi, as;
function Qp() {
  return as || (as = 1, Bi = (e, t = process.argv) => {
    const n = e.startsWith("-") ? "" : e.length === 1 ? "-" : "--", r = t.indexOf(n + e), i = t.indexOf("--");
    return r !== -1 && (i === -1 || r < i);
  }), Bi;
}
var Hi, ss;
function Zp() {
  if (ss) return Hi;
  ss = 1;
  const e = ti, t = Ml, n = Qp(), { env: r } = process;
  let i;
  n("no-color") || n("no-colors") || n("color=false") || n("color=never") ? i = 0 : (n("color") || n("colors") || n("color=true") || n("color=always")) && (i = 1), "FORCE_COLOR" in r && (r.FORCE_COLOR === "true" ? i = 1 : r.FORCE_COLOR === "false" ? i = 0 : i = r.FORCE_COLOR.length === 0 ? 1 : Math.min(parseInt(r.FORCE_COLOR, 10), 3));
  function o(c) {
    return c === 0 ? !1 : {
      level: c,
      hasBasic: !0,
      has256: c >= 2,
      has16m: c >= 3
    };
  }
  function a(c, p) {
    if (i === 0)
      return 0;
    if (n("color=16m") || n("color=full") || n("color=truecolor"))
      return 3;
    if (n("color=256"))
      return 2;
    if (c && !p && i === void 0)
      return 0;
    const l = i || 0;
    if (r.TERM === "dumb")
      return l;
    if (process.platform === "win32") {
      const f = e.release().split(".");
      return Number(f[0]) >= 10 && Number(f[2]) >= 10586 ? Number(f[2]) >= 14931 ? 3 : 2 : 1;
    }
    if ("CI" in r)
      return ["TRAVIS", "CIRCLECI", "APPVEYOR", "GITLAB_CI", "GITHUB_ACTIONS", "BUILDKITE"].some((f) => f in r) || r.CI_NAME === "codeship" ? 1 : l;
    if ("TEAMCITY_VERSION" in r)
      return /^(9\.(0*[1-9]\d*)\.|\d{2,}\.)/.test(r.TEAMCITY_VERSION) ? 1 : 0;
    if (r.COLORTERM === "truecolor")
      return 3;
    if ("TERM_PROGRAM" in r) {
      const f = parseInt((r.TERM_PROGRAM_VERSION || "").split(".")[0], 10);
      switch (r.TERM_PROGRAM) {
        case "iTerm.app":
          return f >= 3 ? 3 : 2;
        case "Apple_Terminal":
          return 2;
      }
    }
    return /-256(color)?$/i.test(r.TERM) ? 2 : /^screen|^xterm|^vt100|^vt220|^rxvt|color|ansi|cygwin|linux/i.test(r.TERM) || "COLORTERM" in r ? 1 : l;
  }
  function s(c) {
    const p = a(c, c && c.isTTY);
    return o(p);
  }
  return Hi = {
    supportsColor: s,
    stdout: o(a(!0, t.isatty(1))),
    stderr: o(a(!0, t.isatty(2)))
  }, Hi;
}
var ls;
function em() {
  return ls || (ls = 1, function(e, t) {
    const n = Ml, r = Po;
    t.init = l, t.log = s, t.formatArgs = o, t.save = c, t.load = p, t.useColors = i, t.destroy = r.deprecate(
      () => {
      },
      "Instance method `debug.destroy()` is deprecated and no longer does anything. It will be removed in the next major version of `debug`."
    ), t.colors = [6, 2, 3, 4, 5, 1];
    try {
      const h = Zp();
      h && (h.stderr || h).level >= 2 && (t.colors = [
        20,
        21,
        26,
        27,
        32,
        33,
        38,
        39,
        40,
        41,
        42,
        43,
        44,
        45,
        56,
        57,
        62,
        63,
        68,
        69,
        74,
        75,
        76,
        77,
        78,
        79,
        80,
        81,
        92,
        93,
        98,
        99,
        112,
        113,
        128,
        129,
        134,
        135,
        148,
        149,
        160,
        161,
        162,
        163,
        164,
        165,
        166,
        167,
        168,
        169,
        170,
        171,
        172,
        173,
        178,
        179,
        184,
        185,
        196,
        197,
        198,
        199,
        200,
        201,
        202,
        203,
        204,
        205,
        206,
        207,
        208,
        209,
        214,
        215,
        220,
        221
      ]);
    } catch {
    }
    t.inspectOpts = Object.keys(process.env).filter((h) => /^debug_/i.test(h)).reduce((h, g) => {
      const _ = g.substring(6).toLowerCase().replace(/_([a-z])/g, (T, A) => A.toUpperCase());
      let E = process.env[g];
      return /^(yes|on|true|enabled)$/i.test(E) ? E = !0 : /^(no|off|false|disabled)$/i.test(E) ? E = !1 : E === "null" ? E = null : E = Number(E), h[_] = E, h;
    }, {});
    function i() {
      return "colors" in t.inspectOpts ? !!t.inspectOpts.colors : n.isatty(process.stderr.fd);
    }
    function o(h) {
      const { namespace: g, useColors: _ } = this;
      if (_) {
        const E = this.color, T = "\x1B[3" + (E < 8 ? E : "8;5;" + E), A = `  ${T};1m${g} \x1B[0m`;
        h[0] = A + h[0].split(`
`).join(`
` + A), h.push(T + "m+" + e.exports.humanize(this.diff) + "\x1B[0m");
      } else
        h[0] = a() + g + " " + h[0];
    }
    function a() {
      return t.inspectOpts.hideDate ? "" : (/* @__PURE__ */ new Date()).toISOString() + " ";
    }
    function s(...h) {
      return process.stderr.write(r.formatWithOptions(t.inspectOpts, ...h) + `
`);
    }
    function c(h) {
      h ? process.env.DEBUG = h : delete process.env.DEBUG;
    }
    function p() {
      return process.env.DEBUG;
    }
    function l(h) {
      h.inspectOpts = {};
      const g = Object.keys(t.inspectOpts);
      for (let _ = 0; _ < g.length; _++)
        h.inspectOpts[g[_]] = t.inspectOpts[g[_]];
    }
    e.exports = Oc()(t);
    const { formatters: f } = e.exports;
    f.o = function(h) {
      return this.inspectOpts.colors = this.useColors, r.inspect(h, this.inspectOpts).split(`
`).map((g) => g.trim()).join(" ");
    }, f.O = function(h) {
      return this.inspectOpts.colors = this.useColors, r.inspect(h, this.inspectOpts);
    };
  }(Cr, Cr.exports)), Cr.exports;
}
typeof process > "u" || process.type === "renderer" || process.browser === !0 || process.__nwjs ? wo.exports = Jp() : wo.exports = em();
var tm = wo.exports, nr = {};
Object.defineProperty(nr, "__esModule", { value: !0 });
nr.ProgressCallbackTransform = void 0;
const nm = Zn;
class rm extends nm.Transform {
  constructor(t, n, r) {
    super(), this.total = t, this.cancellationToken = n, this.onProgress = r, this.start = Date.now(), this.transferred = 0, this.delta = 0, this.nextUpdate = this.start + 1e3;
  }
  _transform(t, n, r) {
    if (this.cancellationToken.cancelled) {
      r(new Error("cancelled"), null);
      return;
    }
    this.transferred += t.length, this.delta += t.length;
    const i = Date.now();
    i >= this.nextUpdate && this.transferred !== this.total && (this.nextUpdate = i + 1e3, this.onProgress({
      total: this.total,
      delta: this.delta,
      transferred: this.transferred,
      percent: this.transferred / this.total * 100,
      bytesPerSecond: Math.round(this.transferred / ((i - this.start) / 1e3))
    }), this.delta = 0), r(null, t);
  }
  _flush(t) {
    if (this.cancellationToken.cancelled) {
      t(new Error("cancelled"));
      return;
    }
    this.onProgress({
      total: this.total,
      delta: this.delta,
      transferred: this.total,
      percent: 100,
      bytesPerSecond: Math.round(this.transferred / ((Date.now() - this.start) / 1e3))
    }), this.delta = 0, t(null);
  }
}
nr.ProgressCallbackTransform = rm;
Object.defineProperty(de, "__esModule", { value: !0 });
de.DigestTransform = de.HttpExecutor = de.HttpError = void 0;
de.addSensitiveRedirectHeader = um;
de.addSensitiveFieldPattern = fm;
de.createHttpError = To;
de.parseJson = hm;
de.configureRequestOptionsFromUrl = Fc;
de.configureRequestUrl = Wo;
de.safeGetHeader = un;
de.configureRequestOptions = Wr;
de.isSensitiveFieldName = xc;
de.hashSensitiveValue = Uc;
de.safeStringifyJson = rn;
const Ic = er, im = tm, om = Ct, am = Zn, vo = bt, sm = Tt, cs = gn, lm = nr, mt = (0, im.default)("electron-builder"), Go = (e) => e.toLowerCase().replace(/[-_]/g, ""), Dc = /* @__PURE__ */ new Set(["authorization", "proxyauthorization", "privatetoken", "xapikey", "xauthtoken", "xaccesstoken", "xgitlabtoken", "cookie", "xcsrftoken"]), Pc = ["token", "password", "secret", "authorization", "credential", "apikey", "passphrase", "auth"], cm = ["key"];
function um(e) {
  Dc.add(Go(e));
}
function fm(e) {
  Pc.push(e.toLowerCase().replace(/[-_]/g, ""));
}
function To(e, t = null) {
  return new qo(e.statusCode || -1, `${e.statusCode} ${e.statusMessage}` + (t == null ? "" : `
` + JSON.stringify(t, null, "  ")) + `
Headers: ` + rn(e.headers), t);
}
const dm = /* @__PURE__ */ new Map([
  [429, "Too many requests"],
  [400, "Bad request"],
  [403, "Forbidden"],
  [404, "Not found"],
  [405, "Method not allowed"],
  [406, "Not acceptable"],
  [408, "Request timeout"],
  [413, "Request entity too large"],
  [500, "Internal server error"],
  [502, "Bad gateway"],
  [503, "Service unavailable"],
  [504, "Gateway timeout"],
  [505, "HTTP version not supported"]
]);
class qo extends Error {
  constructor(t, n = `HTTP error: ${dm.get(t) || t}`, r = null) {
    super(n), this.statusCode = t, this.description = r, this.name = "HttpError", this.code = `HTTP_ERROR_${t}`;
  }
  isServerError() {
    return this.statusCode >= 500 && this.statusCode <= 599;
  }
}
de.HttpError = qo;
function hm(e) {
  return e.then((t) => t == null || t.length === 0 ? null : JSON.parse(t));
}
class nn {
  constructor() {
    this.maxRedirects = 10;
  }
  request(t, n = new sm.CancellationToken(), r) {
    Wr(t);
    const i = r == null ? void 0 : JSON.stringify(r), o = i ? Buffer.from(i) : void 0;
    if (o != null) {
      mt.enabled && mt(rn(r));
      const { headers: a, ...s } = t;
      t = {
        method: "post",
        headers: {
          "Content-Type": "application/json",
          "Content-Length": o.length,
          ...a
        },
        ...s
      };
    }
    return this.doApiRequest(t, n, (a) => a.end(o));
  }
  doApiRequest(t, n, r, i = 0) {
    if (mt.enabled) {
      const { headers: o, auth: a, ...s } = t;
      mt(`Request: ${rn(s)}`);
    }
    return n.createPromise((o, a, s) => {
      const c = this.createRequest(t, (p) => {
        try {
          this.handleResponse(p, t, n, o, a, i, r);
        } catch (l) {
          a(l);
        }
      });
      this.addErrorAndTimeoutHandlers(c, a, t.timeout), this.addRedirectHandlers(c, t, a, i, (p) => {
        this.doApiRequest(p, n, r, i).then(o).catch(a);
      }), r(c, a), s(() => c.abort());
    });
  }
  // noinspection JSUnusedLocalSymbols
  // eslint-disable-next-line
  addRedirectHandlers(t, n, r, i, o) {
  }
  addErrorAndTimeoutHandlers(t, n, r = 60 * 1e3) {
    this.addTimeOutHandler(t, n, r), t.on("error", n), t.on("aborted", () => {
      n(new Error("Request has been aborted by the server"));
    });
  }
  handleResponse(t, n, r, i, o, a, s) {
    var c;
    if (mt.enabled) {
      const { headers: g, auth: _, ...E } = n;
      mt(`Response: ${t.statusCode} ${t.statusMessage}, request options: ${rn(E)}`);
    }
    if (t.statusCode === 404) {
      o(To(t, `method: ${n.method || "GET"} url: ${n.protocol || "https:"}//${n.hostname}${n.port ? `:${n.port}` : ""}${n.path}

Please double check that your authentication token is correct. Due to security reasons, actual status maybe not reported, but 404.
`));
      return;
    } else if (t.statusCode === 204) {
      i();
      return;
    }
    const p = (c = t.statusCode) !== null && c !== void 0 ? c : 0, l = p >= 300 && p < 400, f = un(t, "location");
    if (l && f != null) {
      if (a > this.maxRedirects) {
        o(this.createMaxRedirectError());
        return;
      }
      this.doApiRequest(nn.prepareRedirectUrlOptions(f, n), r, s, a).then(i).catch(o);
      return;
    }
    t.setEncoding("utf8");
    let h = "";
    t.on("error", o), t.on("data", (g) => h += g), t.on("end", () => {
      try {
        if (t.statusCode != null && t.statusCode >= 400) {
          const g = un(t, "content-type"), _ = g != null && (Array.isArray(g) ? g.find((E) => E.includes("json")) != null : g.includes("json"));
          o(To(t, `method: ${n.method || "GET"} url: ${n.protocol || "https:"}//${n.hostname}${n.port ? `:${n.port}` : ""}${n.path}

          Data:
          ${_ ? rn(JSON.parse(h)) : h}
          `));
        } else
          i(h.length === 0 ? null : h);
      } catch (g) {
        o(g);
      }
    });
  }
  async downloadToBuffer(t, n) {
    return await n.cancellationToken.createPromise((r, i, o) => {
      const a = [], s = {
        headers: n.headers || void 0,
        // because PrivateGitHubProvider requires HttpExecutor.prepareRedirectUrlOptions logic, so, we need to redirect manually
        redirect: "manual"
      };
      Wo(t, s), Wr(s), this.doDownload(s, {
        destination: null,
        options: n,
        onCancel: o,
        callback: (c) => {
          c == null ? r(Buffer.concat(a)) : i(c);
        },
        responseHandler: (c, p) => {
          let l = 0;
          c.on("data", (f) => {
            if (l += f.length, l > 524288e3) {
              p(new Error("Maximum allowed size is 500 MB"));
              return;
            }
            a.push(f);
          }), c.on("end", () => {
            p(null);
          });
        }
      }, 0);
    });
  }
  doDownload(t, n, r) {
    const i = this.createRequest(t, (o) => {
      if (o.statusCode >= 400) {
        n.callback(new Error(`Cannot download "${t.protocol || "https:"}//${t.hostname}${t.path}", status ${o.statusCode}: ${o.statusMessage}`));
        return;
      }
      o.on("error", n.callback);
      const a = un(o, "location");
      if (a != null) {
        r < this.maxRedirects ? this.doDownload(nn.prepareRedirectUrlOptions(a, t), n, r++) : n.callback(this.createMaxRedirectError());
        return;
      }
      n.responseHandler == null ? mm(n, o) : n.responseHandler(o, n.callback);
    });
    this.addErrorAndTimeoutHandlers(i, n.callback, t.timeout), this.addRedirectHandlers(i, t, n.callback, r, (o) => {
      this.doDownload(o, n, r++);
    }), i.end();
  }
  createMaxRedirectError() {
    return new Error(`Too many redirects (> ${this.maxRedirects})`);
  }
  addTimeOutHandler(t, n, r) {
    t.on("socket", (i) => {
      i.setTimeout(r, () => {
        t.abort(), n(new Error("Request timed out"));
      });
    });
  }
  static prepareRedirectUrlOptions(t, n) {
    const r = Fc(t, { ...n }), i = r.headers;
    if (i == null)
      return r;
    const o = nn.reconstructOriginalUrl(n), a = Lc(t, n);
    if (nn.isCrossOriginRedirect(o, a)) {
      mt.enabled && mt(`Cross-origin redirect (${o.host} → ${a.host}): stripping sensitive headers`);
      for (const s of Object.keys(i))
        Dc.has(Go(s)) && delete i[s];
    }
    return r;
  }
  static reconstructOriginalUrl(t) {
    const n = t.protocol || "https:";
    if (!t.hostname)
      throw new Error("Missing hostname in request options");
    const r = t.hostname, i = t.port ? `:${t.port}` : "", o = t.path || "/";
    return new vo.URL(`${n}//${r}${i}${o}`);
  }
  static isCrossOriginRedirect(t, n) {
    if (t.hostname.toLowerCase() !== n.hostname.toLowerCase())
      return !0;
    if (t.protocol === "http:" && // This can be replaced with `!originalUrl.port`, but for the sake of clarity.
    ["80", ""].includes(t.port) && n.protocol === "https:" && // This can be replaced with `!redirectUrl.port`, but for the sake of clarity.
    ["443", ""].includes(n.port))
      return !1;
    if (t.protocol !== n.protocol)
      return !0;
    const r = t.port, i = n.port;
    return r !== i;
  }
  static async retryOnServerError(t, n = 3) {
    for (let r = 0; ; r++)
      try {
        return await t();
      } catch (i) {
        if (r < n && (i instanceof qo && i.isServerError() || i.code === "EPIPE")) {
          await new Promise((o) => setTimeout(o, 1e3 * (r + 1)));
          continue;
        }
        throw i;
      }
  }
}
de.HttpExecutor = nn;
function Lc(e, t) {
  try {
    return new vo.URL(e);
  } catch {
    const n = t.hostname, r = t.protocol || "https:", i = t.port ? `:${t.port}` : "", o = `${r}//${n}${i}`;
    return new vo.URL(e, o);
  }
}
function Fc(e, t) {
  const n = Wr(t), r = Lc(e, t);
  return Wo(r, n), n;
}
function Wo(e, t) {
  t.protocol = e.protocol, t.hostname = e.hostname, e.port ? t.port = e.port : t.port && delete t.port, t.path = e.pathname + e.search;
}
class Ao extends am.Transform {
  // noinspection JSUnusedGlobalSymbols
  get actual() {
    return this._actual;
  }
  constructor(t, n = "sha512", r = "base64") {
    super(), this.expected = t, this.algorithm = n, this.encoding = r, this._actual = null, this.isValidateOnEnd = !0, this.digester = (0, Ic.createHash)(n);
  }
  // noinspection JSUnusedGlobalSymbols
  _transform(t, n, r) {
    this.digester.update(t), r(null, t);
  }
  // noinspection JSUnusedGlobalSymbols
  _flush(t) {
    if (this._actual = this.digester.digest(this.encoding), this.isValidateOnEnd)
      try {
        this.validate();
      } catch (n) {
        t(n);
        return;
      }
    t(null);
  }
  validate() {
    if (this._actual == null)
      throw (0, cs.newError)("Not finished yet", "ERR_STREAM_NOT_FINISHED");
    if (this._actual !== this.expected)
      throw (0, cs.newError)(`${this.algorithm} checksum mismatch, expected ${this.expected}, got ${this._actual}`, "ERR_CHECKSUM_MISMATCH");
    return null;
  }
}
de.DigestTransform = Ao;
function pm(e, t, n) {
  return e != null && t != null && e !== t ? (n(new Error(`checksum mismatch: expected ${t} but got ${e} (X-Checksum-Sha2 header)`)), !1) : !0;
}
function un(e, t) {
  const n = e.headers[t];
  return n == null ? null : Array.isArray(n) ? n.length === 0 ? null : n[n.length - 1] : n;
}
function mm(e, t) {
  if (!pm(un(t, "X-Checksum-Sha2"), e.options.sha2, e.callback))
    return;
  const n = [];
  if (e.options.onProgress != null) {
    const a = un(t, "content-length");
    a != null && n.push(new lm.ProgressCallbackTransform(parseInt(a, 10), e.options.cancellationToken, e.options.onProgress));
  }
  const r = e.options.sha512;
  r != null ? n.push(new Ao(r, "sha512", r.length === 128 && !r.includes("+") && !r.includes("Z") && !r.includes("=") ? "hex" : "base64")) : e.options.sha2 != null && n.push(new Ao(e.options.sha2, "sha256", "hex"));
  const i = (0, om.createWriteStream)(e.destination);
  n.push(i);
  let o = t;
  for (const a of n)
    a.on("error", (s) => {
      i.close(), e.options.cancellationToken.cancelled || e.callback(s);
    }), o = o.pipe(a);
  i.on("finish", () => {
    i.close(e.callback);
  });
}
function Wr(e, t, n) {
  n != null && (e.method = n), e.headers = { ...e.headers };
  const r = e.headers;
  return t != null && (r.authorization = t.startsWith("Basic") || t.startsWith("Bearer") ? t : `token ${t}`), r["User-Agent"] == null && (r["User-Agent"] = "electron-builder"), (n == null || n === "GET" || r["Cache-Control"] == null) && (r["Cache-Control"] = "no-cache"), e.protocol == null && process.versions.electron != null && (e.protocol = "https:"), e;
}
function xc(e) {
  const t = Go(e);
  return Pc.some((n) => t.includes(n)) || cm.some((n) => t.endsWith(n));
}
function Uc(e) {
  return `${(0, Ic.createHash)("sha256").update(e).digest("hex")} (sha256 hash)`;
}
function rn(e, t) {
  return JSON.stringify(e, (n, r) => xc(n) || t != null && t.has(n) ? typeof r == "string" ? Uc(r) : "<stripped sensitive data>" : r, 2);
}
var ai = {};
Object.defineProperty(ai, "__esModule", { value: !0 });
ai.MemoLazy = void 0;
class gm {
  constructor(t, n) {
    this.selector = t, this.creator = n, this.selected = void 0, this._value = void 0;
  }
  get hasValue() {
    return this._value !== void 0;
  }
  get value() {
    const t = this.selector();
    if (this._value !== void 0 && kc(this.selected, t))
      return this._value;
    this.selected = t;
    const n = this.creator(t);
    return this.value = n, n;
  }
  set value(t) {
    this._value = t;
  }
}
ai.MemoLazy = gm;
function kc(e, t) {
  if (typeof e == "object" && e !== null && (typeof t == "object" && t !== null)) {
    const i = Object.keys(e), o = Object.keys(t);
    return i.length === o.length && i.every((a) => kc(e[a], t[a]));
  }
  return e === t;
}
var rr = {};
Object.defineProperty(rr, "__esModule", { value: !0 });
rr.githubUrl = Em;
rr.githubTagPrefix = ym;
rr.getS3LikeProviderBaseUrl = _m;
function Em(e, t = "github.com") {
  return `${e.protocol || "https"}://${e.host || t}`;
}
function ym(e) {
  var t;
  return e.tagNamePrefix ? e.tagNamePrefix : !((t = e.vPrefixedTagName) !== null && t !== void 0) || t ? "v" : "";
}
function _m(e) {
  const t = e.provider;
  if (t === "s3")
    return wm(e);
  if (t === "spaces")
    return vm(e);
  throw new Error(`Not supported provider: ${t}`);
}
function wm(e) {
  let t;
  if (e.accelerate == !0)
    t = `https://${e.bucket}.s3-accelerate.amazonaws.com`;
  else if (e.endpoint != null)
    t = `${e.endpoint}/${e.bucket}`;
  else if (e.bucket.includes(".")) {
    if (e.region == null)
      throw new Error(`Bucket name "${e.bucket}" includes a dot, but S3 region is missing`);
    e.region === "us-east-1" ? t = `https://s3.amazonaws.com/${e.bucket}` : t = `https://s3-${e.region}.amazonaws.com/${e.bucket}`;
  } else e.region === "cn-north-1" ? t = `https://${e.bucket}.s3.${e.region}.amazonaws.com.cn` : t = `https://${e.bucket}.s3.amazonaws.com`;
  return Mc(t, e.path);
}
function Mc(e, t) {
  return t != null && t.length > 0 && (t.startsWith("/") || (e += "/"), e += t), e;
}
function vm(e) {
  if (e.name == null)
    throw new Error("name is missing");
  if (e.region == null)
    throw new Error("region is missing");
  return Mc(`https://${e.name}.${e.region}.digitaloceanspaces.com`, e.path);
}
var Vo = {};
Object.defineProperty(Vo, "__esModule", { value: !0 });
Vo.retry = Bc;
const Tm = Tt;
async function Bc(e, t) {
  var n;
  const { retries: r, interval: i, backoff: o = 0, attempt: a = 0, shouldRetry: s, cancellationToken: c = new Tm.CancellationToken() } = t;
  try {
    return await e();
  } catch (p) {
    if (await Promise.resolve((n = s == null ? void 0 : s(p)) !== null && n !== void 0 ? n : !0) && r > 0 && !c.cancelled)
      return await new Promise((l) => setTimeout(l, i + o * a)), await Bc(e, { ...t, retries: r - 1, attempt: a + 1 });
    throw p;
  }
}
var zo = {};
Object.defineProperty(zo, "__esModule", { value: !0 });
zo.parseDn = Am;
function Am(e) {
  let t = !1, n = null, r = "", i = 0;
  e = e.trim();
  const o = /* @__PURE__ */ new Map();
  for (let a = 0; a <= e.length; a++) {
    if (a === e.length) {
      n !== null && o.set(n, r);
      break;
    }
    const s = e[a];
    if (t) {
      if (s === '"') {
        t = !1;
        continue;
      }
    } else {
      if (s === '"') {
        t = !0;
        continue;
      }
      if (s === "\\") {
        a++;
        const c = parseInt(e.slice(a, a + 2), 16);
        Number.isNaN(c) ? r += e[a] : (a++, r += String.fromCharCode(c));
        continue;
      }
      if (n === null && s === "=") {
        n = r, r = "";
        continue;
      }
      if (s === "," || s === ";" || s === "+") {
        n !== null && o.set(n, r), n = null, r = "";
        continue;
      }
    }
    if (s === " " && !t) {
      if (r.length === 0)
        continue;
      if (a > i) {
        let c = a;
        for (; e[c] === " "; )
          c++;
        i = c;
      }
      if (i >= e.length || e[i] === "," || e[i] === ";" || n === null && e[i] === "=" || n !== null && e[i] === "+") {
        a = i - 1;
        continue;
      }
    }
    r += s;
  }
  return o;
}
var hn = {};
Object.defineProperty(hn, "__esModule", { value: !0 });
hn.nil = hn.UUID = void 0;
const Hc = er, jc = gn, Sm = "options.name must be either a string or a Buffer", us = (0, Hc.randomBytes)(16);
us[0] = us[0] | 1;
const Hr = {}, V = [];
for (let e = 0; e < 256; e++) {
  const t = (e + 256).toString(16).substr(1);
  Hr[t] = e, V[e] = t;
}
class jt {
  constructor(t) {
    this.ascii = null, this.binary = null;
    const n = jt.check(t);
    if (!n)
      throw new Error("not a UUID");
    this.version = n.version, n.format === "ascii" ? this.ascii = t : this.binary = t;
  }
  static v5(t, n) {
    return Cm(t, "sha1", 80, n);
  }
  toString() {
    return this.ascii == null && (this.ascii = bm(this.binary)), this.ascii;
  }
  inspect() {
    return `UUID v${this.version} ${this.toString()}`;
  }
  static check(t, n = 0) {
    if (typeof t == "string")
      return t = t.toLowerCase(), /^[a-f0-9]{8}(-[a-f0-9]{4}){3}-([a-f0-9]{12})$/.test(t) ? t === "00000000-0000-0000-0000-000000000000" ? { version: void 0, variant: "nil", format: "ascii" } : {
        version: (Hr[t[14] + t[15]] & 240) >> 4,
        variant: fs((Hr[t[19] + t[20]] & 224) >> 5),
        format: "ascii"
      } : !1;
    if (Buffer.isBuffer(t)) {
      if (t.length < n + 16)
        return !1;
      let r = 0;
      for (; r < 16 && t[n + r] === 0; r++)
        ;
      return r === 16 ? { version: void 0, variant: "nil", format: "binary" } : {
        version: (t[n + 6] & 240) >> 4,
        variant: fs((t[n + 8] & 224) >> 5),
        format: "binary"
      };
    }
    throw (0, jc.newError)("Unknown type of uuid", "ERR_UNKNOWN_UUID_TYPE");
  }
  // read stringified uuid into a Buffer
  static parse(t) {
    const n = Buffer.allocUnsafe(16);
    let r = 0;
    for (let i = 0; i < 16; i++)
      n[i] = Hr[t[r++] + t[r++]], (i === 3 || i === 5 || i === 7 || i === 9) && (r += 1);
    return n;
  }
}
hn.UUID = jt;
jt.OID = jt.parse("6ba7b812-9dad-11d1-80b4-00c04fd430c8");
function fs(e) {
  switch (e) {
    case 0:
    case 1:
    case 3:
      return "ncs";
    case 4:
    case 5:
      return "rfc4122";
    case 6:
      return "microsoft";
    default:
      return "future";
  }
}
var xn;
(function(e) {
  e[e.ASCII = 0] = "ASCII", e[e.BINARY = 1] = "BINARY", e[e.OBJECT = 2] = "OBJECT";
})(xn || (xn = {}));
function Cm(e, t, n, r, i = xn.ASCII) {
  const o = (0, Hc.createHash)(t);
  if (typeof e != "string" && !Buffer.isBuffer(e))
    throw (0, jc.newError)(Sm, "ERR_INVALID_UUID_NAME");
  o.update(r), o.update(e);
  const s = o.digest();
  let c;
  switch (i) {
    case xn.BINARY:
      s[6] = s[6] & 15 | n, s[8] = s[8] & 63 | 128, c = s;
      break;
    case xn.OBJECT:
      s[6] = s[6] & 15 | n, s[8] = s[8] & 63 | 128, c = new jt(s);
      break;
    default:
      c = V[s[0]] + V[s[1]] + V[s[2]] + V[s[3]] + "-" + V[s[4]] + V[s[5]] + "-" + V[s[6] & 15 | n] + V[s[7]] + "-" + V[s[8] & 63 | 128] + V[s[9]] + "-" + V[s[10]] + V[s[11]] + V[s[12]] + V[s[13]] + V[s[14]] + V[s[15]];
      break;
  }
  return c;
}
function bm(e) {
  return V[e[0]] + V[e[1]] + V[e[2]] + V[e[3]] + "-" + V[e[4]] + V[e[5]] + "-" + V[e[6]] + V[e[7]] + "-" + V[e[8]] + V[e[9]] + "-" + V[e[10]] + V[e[11]] + V[e[12]] + V[e[13]] + V[e[14]] + V[e[15]];
}
hn.nil = new jt("00000000-0000-0000-0000-000000000000");
var ir = {}, Gc = {};
(function(e) {
  (function(t) {
    t.parser = function(d, u) {
      return new r(d, u);
    }, t.SAXParser = r, t.SAXStream = f, t.createStream = p, t.MAX_BUFFER_LENGTH = 64 * 1024;
    var n = [
      "comment",
      "sgmlDecl",
      "textNode",
      "tagName",
      "doctype",
      "procInstName",
      "procInstBody",
      "entity",
      "attribName",
      "attribValue",
      "cdata",
      "script"
    ];
    t.EVENTS = [
      "text",
      "processinginstruction",
      "sgmldeclaration",
      "doctype",
      "comment",
      "opentagstart",
      "attribute",
      "opentag",
      "closetag",
      "opencdata",
      "cdata",
      "closecdata",
      "error",
      "end",
      "ready",
      "script",
      "opennamespace",
      "closenamespace"
    ];
    function r(d, u) {
      if (!(this instanceof r))
        return new r(d, u);
      var S = this;
      o(S), S.q = S.c = "", S.bufferCheckPosition = t.MAX_BUFFER_LENGTH, S.encoding = null, S.opt = u || {}, S.opt.lowercase = S.opt.lowercase || S.opt.lowercasetags, S.looseCase = S.opt.lowercase ? "toLowerCase" : "toUpperCase", S.opt.maxEntityCount = S.opt.maxEntityCount || 512, S.opt.maxEntityDepth = S.opt.maxEntityDepth || 4, S.entityCount = S.entityDepth = 0, S.tags = [], S.closed = S.closedRoot = S.sawRoot = !1, S.tag = S.error = null, S.strict = !!d, S.noscript = !!(d || S.opt.noscript), S.state = y.BEGIN, S.strictEntities = S.opt.strictEntities, S.ENTITIES = S.strictEntities ? Object.create(t.XML_ENTITIES) : Object.create(t.ENTITIES), S.attribList = [], S.opt.xmlns && (S.ns = Object.create(T)), S.opt.unquotedAttributeValues === void 0 && (S.opt.unquotedAttributeValues = !d), S.trackPosition = S.opt.position !== !1, S.trackPosition && (S.position = S.line = S.column = 0), Y(S, "onready");
    }
    Object.create || (Object.create = function(d) {
      function u() {
      }
      u.prototype = d;
      var S = new u();
      return S;
    }), Object.keys || (Object.keys = function(d) {
      var u = [];
      for (var S in d) d.hasOwnProperty(S) && u.push(S);
      return u;
    });
    function i(d) {
      for (var u = Math.max(t.MAX_BUFFER_LENGTH, 10), S = 0, w = 0, z = n.length; w < z; w++) {
        var ne = d[n[w]].length;
        if (ne > u)
          switch (n[w]) {
            case "textNode":
              D(d);
              break;
            case "cdata":
              b(d, "oncdata", d.cdata), d.cdata = "";
              break;
            case "script":
              b(d, "onscript", d.script), d.script = "";
              break;
            default:
              k(d, "Max buffer length exceeded: " + n[w]);
          }
        S = Math.max(S, ne);
      }
      var se = t.MAX_BUFFER_LENGTH - S;
      d.bufferCheckPosition = se + d.position;
    }
    function o(d) {
      for (var u = 0, S = n.length; u < S; u++)
        d[n[u]] = "";
    }
    function a(d) {
      D(d), d.cdata !== "" && (b(d, "oncdata", d.cdata), d.cdata = ""), d.script !== "" && (b(d, "onscript", d.script), d.script = "");
    }
    r.prototype = {
      end: function() {
        W(this);
      },
      write: dr,
      resume: function() {
        return this.error = null, this;
      },
      close: function() {
        return this.write(null);
      },
      flush: function() {
        a(this);
      }
    };
    var s;
    try {
      s = require("stream").Stream;
    } catch {
      s = function() {
      };
    }
    s || (s = function() {
    });
    var c = t.EVENTS.filter(function(d) {
      return d !== "error" && d !== "end";
    });
    function p(d, u) {
      return new f(d, u);
    }
    function l(d, u) {
      if (d.length >= 2) {
        if (d[0] === 255 && d[1] === 254)
          return "utf-16le";
        if (d[0] === 254 && d[1] === 255)
          return "utf-16be";
      }
      return d.length >= 3 && d[0] === 239 && d[1] === 187 && d[2] === 191 ? "utf8" : d.length >= 4 ? d[0] === 60 && d[1] === 0 && d[2] === 63 && d[3] === 0 ? "utf-16le" : d[0] === 0 && d[1] === 60 && d[2] === 0 && d[3] === 63 ? "utf-16be" : "utf8" : u ? "utf8" : null;
    }
    function f(d, u) {
      if (!(this instanceof f))
        return new f(d, u);
      s.apply(this), this._parser = new r(d, u), this.writable = !0, this.readable = !0;
      var S = this;
      this._parser.onend = function() {
        S.emit("end");
      }, this._parser.onerror = function(w) {
        S.emit("error", w), S._parser.error = null;
      }, this._decoder = null, this._decoderBuffer = null, c.forEach(function(w) {
        Object.defineProperty(S, "on" + w, {
          get: function() {
            return S._parser["on" + w];
          },
          set: function(z) {
            if (!z)
              return S.removeAllListeners(w), S._parser["on" + w] = z, z;
            S.on(w, z);
          },
          enumerable: !0,
          configurable: !1
        });
      });
    }
    f.prototype = Object.create(s.prototype, {
      constructor: {
        value: f
      }
    }), f.prototype._decodeBuffer = function(d, u) {
      if (this._decoderBuffer && (d = Buffer.concat([this._decoderBuffer, d]), this._decoderBuffer = null), !this._decoder) {
        var S = l(d, u);
        if (!S)
          return this._decoderBuffer = d, "";
        this._parser.encoding = S, this._decoder = new TextDecoder(S);
      }
      return this._decoder.decode(d, { stream: !u });
    }, f.prototype.write = function(d) {
      if (typeof Buffer == "function" && typeof Buffer.isBuffer == "function" && Buffer.isBuffer(d))
        d = this._decodeBuffer(d, !1);
      else if (this._decoderBuffer) {
        var u = this._decodeBuffer(Buffer.alloc(0), !0);
        u && (this._parser.write(u), this.emit("data", u));
      }
      return this._parser.write(d.toString()), this.emit("data", d), !0;
    }, f.prototype.end = function(d) {
      if (d && d.length && this.write(d), this._decoderBuffer) {
        var u = this._decodeBuffer(Buffer.alloc(0), !0);
        u && (this._parser.write(u), this.emit("data", u));
      } else if (this._decoder) {
        var S = this._decoder.decode();
        S && (this._parser.write(S), this.emit("data", S));
      }
      return this._parser.end(), !0;
    }, f.prototype.on = function(d, u) {
      var S = this;
      return !S._parser["on" + d] && c.indexOf(d) !== -1 && (S._parser["on" + d] = function() {
        var w = arguments.length === 1 ? [arguments[0]] : Array.apply(null, arguments);
        w.splice(0, 0, d), S.emit.apply(S, w);
      }), s.prototype.on.call(S, d, u);
    };
    var h = "[CDATA[", g = "DOCTYPE", _ = "http://www.w3.org/XML/1998/namespace", E = "http://www.w3.org/2000/xmlns/", T = { xml: _, xmlns: E }, A = /[:_A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF\uF900-\uFDCF\uFDF0-\uFFFD]/, R = /[:_A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF\uF900-\uFDCF\uFDF0-\uFFFD\u00B7\u0300-\u036F\u203F-\u2040.\d-]/, P = /[#:_A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF\uF900-\uFDCF\uFDF0-\uFFFD]/, B = /[#:_A-Za-z\u00C0-\u00D6\u00D8-\u00F6\u00F8-\u02FF\u0370-\u037D\u037F-\u1FFF\u200C-\u200D\u2070-\u218F\u2C00-\u2FEF\u3001-\uD7FF\uF900-\uFDCF\uFDF0-\uFFFD\u00B7\u0300-\u036F\u203F-\u2040.\d-]/;
    function q(d) {
      return d === " " || d === `
` || d === "\r" || d === "	";
    }
    function Q(d) {
      return d === '"' || d === "'";
    }
    function Z(d) {
      return d === ">" || q(d);
    }
    function ae(d, u) {
      return d.test(u);
    }
    function U(d, u) {
      return !ae(d, u);
    }
    var y = 0;
    t.STATE = {
      BEGIN: y++,
      // leading byte order mark or whitespace
      BEGIN_WHITESPACE: y++,
      // leading whitespace
      TEXT: y++,
      // general stuff
      TEXT_ENTITY: y++,
      // &amp and such.
      OPEN_WAKA: y++,
      // <
      SGML_DECL: y++,
      // <!BLARG
      SGML_DECL_QUOTED: y++,
      // <!BLARG foo "bar
      DOCTYPE: y++,
      // <!DOCTYPE
      DOCTYPE_QUOTED: y++,
      // <!DOCTYPE "//blah
      DOCTYPE_DTD: y++,
      // <!DOCTYPE "//blah" [ ...
      DOCTYPE_DTD_QUOTED: y++,
      // <!DOCTYPE "//blah" [ "foo
      COMMENT_STARTING: y++,
      // <!-
      COMMENT: y++,
      // <!--
      COMMENT_ENDING: y++,
      // <!-- blah -
      COMMENT_ENDED: y++,
      // <!-- blah --
      CDATA: y++,
      // <![CDATA[ something
      CDATA_ENDING: y++,
      // ]
      CDATA_ENDING_2: y++,
      // ]]
      PROC_INST: y++,
      // <?hi
      PROC_INST_BODY: y++,
      // <?hi there
      PROC_INST_ENDING: y++,
      // <?hi "there" ?
      OPEN_TAG: y++,
      // <strong
      OPEN_TAG_SLASH: y++,
      // <strong /
      ATTRIB: y++,
      // <a
      ATTRIB_NAME: y++,
      // <a foo
      ATTRIB_NAME_SAW_WHITE: y++,
      // <a foo _
      ATTRIB_VALUE: y++,
      // <a foo=
      ATTRIB_VALUE_QUOTED: y++,
      // <a foo="bar
      ATTRIB_VALUE_CLOSED: y++,
      // <a foo="bar"
      ATTRIB_VALUE_UNQUOTED: y++,
      // <a foo=bar
      ATTRIB_VALUE_ENTITY_Q: y++,
      // <foo bar="&quot;"
      ATTRIB_VALUE_ENTITY_U: y++,
      // <foo bar=&quot
      CLOSE_TAG: y++,
      // </a
      CLOSE_TAG_SAW_WHITE: y++,
      // </a   >
      SCRIPT: y++,
      // <script> ...
      SCRIPT_ENDING: y++
      // <script> ... <
    }, t.XML_ENTITIES = {
      amp: "&",
      gt: ">",
      lt: "<",
      quot: '"',
      apos: "'"
    }, t.ENTITIES = {
      amp: "&",
      gt: ">",
      lt: "<",
      quot: '"',
      apos: "'",
      AElig: 198,
      Aacute: 193,
      Acirc: 194,
      Agrave: 192,
      Aring: 197,
      Atilde: 195,
      Auml: 196,
      Ccedil: 199,
      ETH: 208,
      Eacute: 201,
      Ecirc: 202,
      Egrave: 200,
      Euml: 203,
      Iacute: 205,
      Icirc: 206,
      Igrave: 204,
      Iuml: 207,
      Ntilde: 209,
      Oacute: 211,
      Ocirc: 212,
      Ograve: 210,
      Oslash: 216,
      Otilde: 213,
      Ouml: 214,
      THORN: 222,
      Uacute: 218,
      Ucirc: 219,
      Ugrave: 217,
      Uuml: 220,
      Yacute: 221,
      aacute: 225,
      acirc: 226,
      aelig: 230,
      agrave: 224,
      aring: 229,
      atilde: 227,
      auml: 228,
      ccedil: 231,
      eacute: 233,
      ecirc: 234,
      egrave: 232,
      eth: 240,
      euml: 235,
      iacute: 237,
      icirc: 238,
      igrave: 236,
      iuml: 239,
      ntilde: 241,
      oacute: 243,
      ocirc: 244,
      ograve: 242,
      oslash: 248,
      otilde: 245,
      ouml: 246,
      szlig: 223,
      thorn: 254,
      uacute: 250,
      ucirc: 251,
      ugrave: 249,
      uuml: 252,
      yacute: 253,
      yuml: 255,
      copy: 169,
      reg: 174,
      nbsp: 160,
      iexcl: 161,
      cent: 162,
      pound: 163,
      curren: 164,
      yen: 165,
      brvbar: 166,
      sect: 167,
      uml: 168,
      ordf: 170,
      laquo: 171,
      not: 172,
      shy: 173,
      macr: 175,
      deg: 176,
      plusmn: 177,
      sup1: 185,
      sup2: 178,
      sup3: 179,
      acute: 180,
      micro: 181,
      para: 182,
      middot: 183,
      cedil: 184,
      ordm: 186,
      raquo: 187,
      frac14: 188,
      frac12: 189,
      frac34: 190,
      iquest: 191,
      times: 215,
      divide: 247,
      OElig: 338,
      oelig: 339,
      Scaron: 352,
      scaron: 353,
      Yuml: 376,
      fnof: 402,
      circ: 710,
      tilde: 732,
      Alpha: 913,
      Beta: 914,
      Gamma: 915,
      Delta: 916,
      Epsilon: 917,
      Zeta: 918,
      Eta: 919,
      Theta: 920,
      Iota: 921,
      Kappa: 922,
      Lambda: 923,
      Mu: 924,
      Nu: 925,
      Xi: 926,
      Omicron: 927,
      Pi: 928,
      Rho: 929,
      Sigma: 931,
      Tau: 932,
      Upsilon: 933,
      Phi: 934,
      Chi: 935,
      Psi: 936,
      Omega: 937,
      alpha: 945,
      beta: 946,
      gamma: 947,
      delta: 948,
      epsilon: 949,
      zeta: 950,
      eta: 951,
      theta: 952,
      iota: 953,
      kappa: 954,
      lambda: 955,
      mu: 956,
      nu: 957,
      xi: 958,
      omicron: 959,
      pi: 960,
      rho: 961,
      sigmaf: 962,
      sigma: 963,
      tau: 964,
      upsilon: 965,
      phi: 966,
      chi: 967,
      psi: 968,
      omega: 969,
      thetasym: 977,
      upsih: 978,
      piv: 982,
      ensp: 8194,
      emsp: 8195,
      thinsp: 8201,
      zwnj: 8204,
      zwj: 8205,
      lrm: 8206,
      rlm: 8207,
      ndash: 8211,
      mdash: 8212,
      lsquo: 8216,
      rsquo: 8217,
      sbquo: 8218,
      ldquo: 8220,
      rdquo: 8221,
      bdquo: 8222,
      dagger: 8224,
      Dagger: 8225,
      bull: 8226,
      hellip: 8230,
      permil: 8240,
      prime: 8242,
      Prime: 8243,
      lsaquo: 8249,
      rsaquo: 8250,
      oline: 8254,
      frasl: 8260,
      euro: 8364,
      image: 8465,
      weierp: 8472,
      real: 8476,
      trade: 8482,
      alefsym: 8501,
      larr: 8592,
      uarr: 8593,
      rarr: 8594,
      darr: 8595,
      harr: 8596,
      crarr: 8629,
      lArr: 8656,
      uArr: 8657,
      rArr: 8658,
      dArr: 8659,
      hArr: 8660,
      forall: 8704,
      part: 8706,
      exist: 8707,
      empty: 8709,
      nabla: 8711,
      isin: 8712,
      notin: 8713,
      ni: 8715,
      prod: 8719,
      sum: 8721,
      minus: 8722,
      lowast: 8727,
      radic: 8730,
      prop: 8733,
      infin: 8734,
      ang: 8736,
      and: 8743,
      or: 8744,
      cap: 8745,
      cup: 8746,
      int: 8747,
      there4: 8756,
      sim: 8764,
      cong: 8773,
      asymp: 8776,
      ne: 8800,
      equiv: 8801,
      le: 8804,
      ge: 8805,
      sub: 8834,
      sup: 8835,
      nsub: 8836,
      sube: 8838,
      supe: 8839,
      oplus: 8853,
      otimes: 8855,
      perp: 8869,
      sdot: 8901,
      lceil: 8968,
      rceil: 8969,
      lfloor: 8970,
      rfloor: 8971,
      lang: 9001,
      rang: 9002,
      loz: 9674,
      spades: 9824,
      clubs: 9827,
      hearts: 9829,
      diams: 9830
    }, Object.keys(t.ENTITIES).forEach(function(d) {
      var u = t.ENTITIES[d], S = typeof u == "number" ? String.fromCharCode(u) : u;
      t.ENTITIES[d] = S;
    });
    for (var j in t.STATE)
      t.STATE[t.STATE[j]] = j;
    y = t.STATE;
    function Y(d, u, S) {
      d[u] && d[u](S);
    }
    function te(d) {
      var u = d && d.match(/(?:^|\s)encoding\s*=\s*(['"])([^'"]+)\1/i);
      return u ? u[2] : null;
    }
    function $(d) {
      return d ? d.toLowerCase().replace(/[^a-z0-9]/g, "") : null;
    }
    function N(d, u) {
      const S = $(d), w = $(u);
      return !S || !w ? !0 : w === "utf16" ? S === "utf16le" || S === "utf16be" : S === w;
    }
    function I(d, u) {
      if (!(!d.strict || !d.encoding || !u || u.name !== "xml")) {
        var S = te(u.body);
        S && !N(d.encoding, S) && L(
          d,
          "XML declaration encoding " + S + " does not match detected stream encoding " + d.encoding.toUpperCase()
        );
      }
    }
    function b(d, u, S) {
      d.textNode && D(d), Y(d, u, S);
    }
    function D(d) {
      d.textNode = O(d.opt, d.textNode), d.textNode && Y(d, "ontext", d.textNode), d.textNode = "";
    }
    function O(d, u) {
      return d.trim && (u = u.trim()), d.normalize && (u = u.replace(/\s+/g, " ")), u;
    }
    function k(d, u) {
      return D(d), d.trackPosition && (u += `
Line: ` + d.line + `
Column: ` + d.column + `
Char: ` + d.c), u = new Error(u), d.error = u, Y(d, "onerror", u), d;
    }
    function W(d) {
      return d.sawRoot && !d.closedRoot && L(d, "Unclosed root tag"), d.state !== y.BEGIN && d.state !== y.BEGIN_WHITESPACE && d.state !== y.TEXT && k(d, "Unexpected end"), D(d), d.c = "", d.closed = !0, Y(d, "onend"), r.call(d, d.strict, d.opt), d;
    }
    function L(d, u) {
      if (typeof d != "object" || !(d instanceof r))
        throw new Error("bad call to strictFail");
      d.strict && k(d, u);
    }
    function X(d) {
      d.strict || (d.tagName = d.tagName[d.looseCase]());
      var u = d.tags[d.tags.length - 1] || d, S = d.tag = { name: d.tagName, attributes: {} };
      d.opt.xmlns && (S.ns = u.ns), d.attribList.length = 0, b(d, "onopentagstart", S);
    }
    function ue(d, u) {
      var S = d.indexOf(":"), w = S < 0 ? ["", d] : d.split(":"), z = w[0], ne = w[1];
      return u && d === "xmlns" && (z = "xmlns", ne = ""), { prefix: z, local: ne };
    }
    function M(d) {
      if (d.strict || (d.attribName = d.attribName[d.looseCase]()), d.attribList.indexOf(d.attribName) !== -1 || d.tag.attributes.hasOwnProperty(d.attribName)) {
        d.attribName = d.attribValue = "";
        return;
      }
      if (d.opt.xmlns) {
        var u = ue(d.attribName, !0), S = u.prefix, w = u.local;
        if (S === "xmlns")
          if (w === "xml" && d.attribValue !== _)
            L(
              d,
              "xml: prefix must be bound to " + _ + `
Actual: ` + d.attribValue
            );
          else if (w === "xmlns" && d.attribValue !== E)
            L(
              d,
              "xmlns: prefix must be bound to " + E + `
Actual: ` + d.attribValue
            );
          else {
            var z = d.tag, ne = d.tags[d.tags.length - 1] || d;
            z.ns === ne.ns && (z.ns = Object.create(ne.ns)), z.ns[w] = d.attribValue;
          }
        d.attribList.push([d.attribName, d.attribValue]);
      } else
        d.tag.attributes[d.attribName] = d.attribValue, b(d, "onattribute", {
          name: d.attribName,
          value: d.attribValue
        });
      d.attribName = d.attribValue = "";
    }
    function ve(d, u) {
      if (d.opt.xmlns) {
        var S = d.tag, w = ue(d.tagName);
        S.prefix = w.prefix, S.local = w.local, S.uri = S.ns[w.prefix] || "", S.prefix && !S.uri && (L(
          d,
          "Unbound namespace prefix: " + JSON.stringify(d.tagName)
        ), S.uri = w.prefix);
        var z = d.tags[d.tags.length - 1] || d;
        S.ns && z.ns !== S.ns && Object.keys(S.ns).forEach(function($t) {
          b(d, "onopennamespace", {
            prefix: $t,
            uri: S.ns[$t]
          });
        });
        for (var ne = 0, se = d.attribList.length; ne < se; ne++) {
          var Te = d.attribList[ne], Ae = Te[0], Ge = Te[1], fe = ue(Ae, !0), qe = fe.prefix, Ci = fe.local, hr = qe === "" ? "" : S.ns[qe] || "", Tn = {
            name: Ae,
            value: Ge,
            prefix: qe,
            local: Ci,
            uri: hr
          };
          qe && qe !== "xmlns" && !hr && (L(
            d,
            "Unbound namespace prefix: " + JSON.stringify(qe)
          ), Tn.uri = qe), d.tag.attributes[Ae] = Tn, b(d, "onattribute", Tn);
        }
        d.attribList.length = 0;
      }
      d.tag.isSelfClosing = !!u, d.sawRoot = !0, d.tags.push(d.tag), b(d, "onopentag", d.tag), u || (!d.noscript && d.tagName.toLowerCase() === "script" ? d.state = y.SCRIPT : d.state = y.TEXT, d.tag = null, d.tagName = ""), d.attribName = d.attribValue = "", d.attribList.length = 0;
    }
    function wn(d) {
      if (!d.tagName) {
        L(d, "Weird empty close tag."), d.textNode += "</>", d.state = y.TEXT;
        return;
      }
      if (d.script) {
        if (d.tagName !== "script") {
          d.script += "</" + d.tagName + ">", d.tagName = "", d.state = y.SCRIPT;
          return;
        }
        b(d, "onscript", d.script), d.script = "";
      }
      var u = d.tags.length, S = d.tagName;
      d.strict || (S = S[d.looseCase]());
      for (var w = S; u--; ) {
        var z = d.tags[u];
        if (z.name !== w)
          L(d, "Unexpected close tag");
        else
          break;
      }
      if (u < 0) {
        L(d, "Unmatched closing tag: " + d.tagName), d.textNode += "</" + d.tagName + ">", d.state = y.TEXT;
        return;
      }
      d.tagName = S;
      for (var ne = d.tags.length; ne-- > u; ) {
        var se = d.tag = d.tags.pop();
        d.tagName = d.tag.name, b(d, "onclosetag", d.tagName);
        var Te = {};
        for (var Ae in se.ns)
          Te[Ae] = se.ns[Ae];
        var Ge = d.tags[d.tags.length - 1] || d;
        d.opt.xmlns && se.ns !== Ge.ns && Object.keys(se.ns).forEach(function(fe) {
          var qe = se.ns[fe];
          b(d, "onclosenamespace", { prefix: fe, uri: qe });
        });
      }
      u === 0 && (d.closedRoot = !0), d.tagName = d.attribValue = d.attribName = "", d.attribList.length = 0, d.state = y.TEXT;
    }
    function je(d) {
      var u = d.entity, S = u.toLowerCase(), w, z = "";
      return d.ENTITIES[u] ? d.ENTITIES[u] : d.ENTITIES[S] ? d.ENTITIES[S] : (u = S, u.charAt(0) === "#" && (u.charAt(1) === "x" ? (u = u.slice(2), w = parseInt(u, 16), z = w.toString(16)) : (u = u.slice(1), w = parseInt(u, 10), z = w.toString(10))), u = u.replace(/^0+/, ""), isNaN(w) || z.toLowerCase() !== u || w < 0 || w > 1114111 ? (L(d, "Invalid character entity"), "&" + d.entity + ";") : String.fromCodePoint(w));
    }
    function vn(d, u) {
      u === "<" ? (d.state = y.OPEN_WAKA, d.startTagPosition = d.position) : q(u) || (L(d, "Non-whitespace before first tag."), d.textNode = u, d.state = y.TEXT);
    }
    function Yt(d, u) {
      var S = "";
      return u < d.length && (S = d.charAt(u)), S;
    }
    function dr(d) {
      var u = this;
      if (this.error)
        throw this.error;
      if (u.closed)
        return k(
          u,
          "Cannot write after close. Assign an onready handler."
        );
      if (d === null)
        return W(u);
      typeof d == "object" && (d = d.toString());
      for (var S = 0, w = ""; w = Yt(d, S++), u.c = w, !!w; )
        switch (u.trackPosition && (u.position++, w === `
` ? (u.line++, u.column = 0) : u.column++), u.state) {
          case y.BEGIN:
            if (u.state = y.BEGIN_WHITESPACE, w === "\uFEFF")
              continue;
            vn(u, w);
            continue;
          case y.BEGIN_WHITESPACE:
            vn(u, w);
            continue;
          case y.TEXT:
            if (u.sawRoot && !u.closedRoot) {
              for (var ne = S - 1; w && w !== "<" && w !== "&"; )
                w = Yt(d, S++), w && u.trackPosition && (u.position++, w === `
` ? (u.line++, u.column = 0) : u.column++);
              u.textNode += d.substring(ne, S - 1);
            }
            w === "<" && !(u.sawRoot && u.closedRoot && !u.strict) ? (u.state = y.OPEN_WAKA, u.startTagPosition = u.position) : (!q(w) && (!u.sawRoot || u.closedRoot) && L(u, "Text data outside of root node."), w === "&" ? u.state = y.TEXT_ENTITY : u.textNode += w);
            continue;
          case y.SCRIPT:
            w === "<" ? u.state = y.SCRIPT_ENDING : u.script += w;
            continue;
          case y.SCRIPT_ENDING:
            w === "/" ? u.state = y.CLOSE_TAG : (u.script += "<" + w, u.state = y.SCRIPT);
            continue;
          case y.OPEN_WAKA:
            if (w === "!")
              u.state = y.SGML_DECL, u.sgmlDecl = "";
            else if (!q(w)) if (ae(A, w))
              u.state = y.OPEN_TAG, u.tagName = w;
            else if (w === "/")
              u.state = y.CLOSE_TAG, u.tagName = "";
            else if (w === "?")
              u.state = y.PROC_INST, u.procInstName = u.procInstBody = "";
            else {
              if (L(u, "Unencoded <"), u.startTagPosition + 1 < u.position) {
                var z = u.position - u.startTagPosition;
                w = new Array(z).join(" ") + w;
              }
              u.textNode += "<" + w, u.state = y.TEXT;
            }
            continue;
          case y.SGML_DECL:
            if (u.sgmlDecl + w === "--") {
              u.state = y.COMMENT, u.comment = "", u.sgmlDecl = "";
              continue;
            }
            u.doctype && u.doctype !== !0 && u.sgmlDecl ? (u.state = y.DOCTYPE_DTD, u.doctype += "<!" + u.sgmlDecl + w, u.sgmlDecl = "") : (u.sgmlDecl + w).toUpperCase() === h ? (b(u, "onopencdata"), u.state = y.CDATA, u.sgmlDecl = "", u.cdata = "") : (u.sgmlDecl + w).toUpperCase() === g ? (u.state = y.DOCTYPE, (u.doctype || u.sawRoot) && L(
              u,
              "Inappropriately located doctype declaration"
            ), u.doctype = "", u.sgmlDecl = "") : w === ">" ? (b(u, "onsgmldeclaration", u.sgmlDecl), u.sgmlDecl = "", u.state = y.TEXT) : (Q(w) && (u.state = y.SGML_DECL_QUOTED), u.sgmlDecl += w);
            continue;
          case y.SGML_DECL_QUOTED:
            w === u.q && (u.state = y.SGML_DECL, u.q = ""), u.sgmlDecl += w;
            continue;
          case y.DOCTYPE:
            w === ">" ? (u.state = y.TEXT, b(u, "ondoctype", u.doctype), u.doctype = !0) : (u.doctype += w, w === "[" ? u.state = y.DOCTYPE_DTD : Q(w) && (u.state = y.DOCTYPE_QUOTED, u.q = w));
            continue;
          case y.DOCTYPE_QUOTED:
            u.doctype += w, w === u.q && (u.q = "", u.state = y.DOCTYPE);
            continue;
          case y.DOCTYPE_DTD:
            w === "]" ? (u.doctype += w, u.state = y.DOCTYPE) : w === "<" ? (u.state = y.OPEN_WAKA, u.startTagPosition = u.position) : Q(w) ? (u.doctype += w, u.state = y.DOCTYPE_DTD_QUOTED, u.q = w) : u.doctype += w;
            continue;
          case y.DOCTYPE_DTD_QUOTED:
            u.doctype += w, w === u.q && (u.state = y.DOCTYPE_DTD, u.q = "");
            continue;
          case y.COMMENT:
            w === "-" ? u.state = y.COMMENT_ENDING : u.comment += w;
            continue;
          case y.COMMENT_ENDING:
            w === "-" ? (u.state = y.COMMENT_ENDED, u.comment = O(u.opt, u.comment), u.comment && b(u, "oncomment", u.comment), u.comment = "") : (u.comment += "-" + w, u.state = y.COMMENT);
            continue;
          case y.COMMENT_ENDED:
            w !== ">" ? (L(u, "Malformed comment"), u.comment += "--" + w, u.state = y.COMMENT) : u.doctype && u.doctype !== !0 ? u.state = y.DOCTYPE_DTD : u.state = y.TEXT;
            continue;
          case y.CDATA:
            for (var ne = S - 1; w && w !== "]"; )
              w = Yt(d, S++), w && u.trackPosition && (u.position++, w === `
` ? (u.line++, u.column = 0) : u.column++);
            u.cdata += d.substring(ne, S - 1), w === "]" && (u.state = y.CDATA_ENDING);
            continue;
          case y.CDATA_ENDING:
            w === "]" ? u.state = y.CDATA_ENDING_2 : (u.cdata += "]" + w, u.state = y.CDATA);
            continue;
          case y.CDATA_ENDING_2:
            w === ">" ? (u.cdata && b(u, "oncdata", u.cdata), b(u, "onclosecdata"), u.cdata = "", u.state = y.TEXT) : w === "]" ? u.cdata += "]" : (u.cdata += "]]" + w, u.state = y.CDATA);
            continue;
          case y.PROC_INST:
            w === "?" ? u.state = y.PROC_INST_ENDING : q(w) ? u.state = y.PROC_INST_BODY : u.procInstName += w;
            continue;
          case y.PROC_INST_BODY:
            if (!u.procInstBody && q(w))
              continue;
            w === "?" ? u.state = y.PROC_INST_ENDING : u.procInstBody += w;
            continue;
          case y.PROC_INST_ENDING:
            if (w === ">") {
              const Ge = {
                name: u.procInstName,
                body: u.procInstBody
              };
              I(u, Ge), b(u, "onprocessinginstruction", Ge), u.procInstName = u.procInstBody = "", u.state = y.TEXT;
            } else
              u.procInstBody += "?" + w, u.state = y.PROC_INST_BODY;
            continue;
          case y.OPEN_TAG:
            ae(R, w) ? u.tagName += w : (X(u), w === ">" ? ve(u) : w === "/" ? u.state = y.OPEN_TAG_SLASH : (q(w) || L(u, "Invalid character in tag name"), u.state = y.ATTRIB));
            continue;
          case y.OPEN_TAG_SLASH:
            w === ">" ? (ve(u, !0), wn(u)) : (L(
              u,
              "Forward-slash in opening tag not followed by >"
            ), u.state = y.ATTRIB);
            continue;
          case y.ATTRIB:
            if (q(w))
              continue;
            w === ">" ? ve(u) : w === "/" ? u.state = y.OPEN_TAG_SLASH : ae(A, w) ? (u.attribName = w, u.attribValue = "", u.state = y.ATTRIB_NAME) : L(u, "Invalid attribute name");
            continue;
          case y.ATTRIB_NAME:
            w === "=" ? u.state = y.ATTRIB_VALUE : w === ">" ? (L(u, "Attribute without value"), u.attribValue = u.attribName, M(u), ve(u)) : q(w) ? u.state = y.ATTRIB_NAME_SAW_WHITE : ae(R, w) ? u.attribName += w : L(u, "Invalid attribute name");
            continue;
          case y.ATTRIB_NAME_SAW_WHITE:
            if (w === "=")
              u.state = y.ATTRIB_VALUE;
            else {
              if (q(w))
                continue;
              L(u, "Attribute without value"), u.tag.attributes[u.attribName] = "", u.attribValue = "", b(u, "onattribute", {
                name: u.attribName,
                value: ""
              }), u.attribName = "", w === ">" ? ve(u) : ae(A, w) ? (u.attribName = w, u.state = y.ATTRIB_NAME) : (L(u, "Invalid attribute name"), u.state = y.ATTRIB);
            }
            continue;
          case y.ATTRIB_VALUE:
            if (q(w))
              continue;
            Q(w) ? (u.q = w, u.state = y.ATTRIB_VALUE_QUOTED) : (u.opt.unquotedAttributeValues || k(u, "Unquoted attribute value"), u.state = y.ATTRIB_VALUE_UNQUOTED, u.attribValue = w);
            continue;
          case y.ATTRIB_VALUE_QUOTED:
            if (w !== u.q) {
              w === "&" ? u.state = y.ATTRIB_VALUE_ENTITY_Q : u.attribValue += w;
              continue;
            }
            M(u), u.q = "", u.state = y.ATTRIB_VALUE_CLOSED;
            continue;
          case y.ATTRIB_VALUE_CLOSED:
            q(w) ? u.state = y.ATTRIB : w === ">" ? ve(u) : w === "/" ? u.state = y.OPEN_TAG_SLASH : ae(A, w) ? (L(u, "No whitespace between attributes"), u.attribName = w, u.attribValue = "", u.state = y.ATTRIB_NAME) : L(u, "Invalid attribute name");
            continue;
          case y.ATTRIB_VALUE_UNQUOTED:
            if (!Z(w)) {
              w === "&" ? u.state = y.ATTRIB_VALUE_ENTITY_U : u.attribValue += w;
              continue;
            }
            M(u), w === ">" ? ve(u) : u.state = y.ATTRIB;
            continue;
          case y.CLOSE_TAG:
            if (u.tagName)
              w === ">" ? wn(u) : ae(R, w) ? u.tagName += w : u.script ? (u.script += "</" + u.tagName + w, u.tagName = "", u.state = y.SCRIPT) : (q(w) || L(u, "Invalid tagname in closing tag"), u.state = y.CLOSE_TAG_SAW_WHITE);
            else {
              if (q(w))
                continue;
              U(A, w) ? u.script ? (u.script += "</" + w, u.state = y.SCRIPT) : L(u, "Invalid tagname in closing tag.") : u.tagName = w;
            }
            continue;
          case y.CLOSE_TAG_SAW_WHITE:
            if (q(w))
              continue;
            w === ">" ? wn(u) : L(u, "Invalid characters in closing tag");
            continue;
          case y.TEXT_ENTITY:
          case y.ATTRIB_VALUE_ENTITY_Q:
          case y.ATTRIB_VALUE_ENTITY_U:
            var se, Te;
            switch (u.state) {
              case y.TEXT_ENTITY:
                se = y.TEXT, Te = "textNode";
                break;
              case y.ATTRIB_VALUE_ENTITY_Q:
                se = y.ATTRIB_VALUE_QUOTED, Te = "attribValue";
                break;
              case y.ATTRIB_VALUE_ENTITY_U:
                se = y.ATTRIB_VALUE_UNQUOTED, Te = "attribValue";
                break;
            }
            if (w === ";") {
              var Ae = je(u);
              u.opt.unparsedEntities && !Object.values(t.XML_ENTITIES).includes(Ae) ? ((u.entityCount += 1) > u.opt.maxEntityCount && k(
                u,
                "Parsed entity count exceeds max entity count"
              ), (u.entityDepth += 1) > u.opt.maxEntityDepth && k(
                u,
                "Parsed entity depth exceeds max entity depth"
              ), u.entity = "", u.state = se, u.write(Ae), u.entityDepth -= 1) : (u[Te] += Ae, u.entity = "", u.state = se);
            } else ae(u.entity.length ? B : P, w) ? u.entity += w : (L(u, "Invalid character in entity name"), u[Te] += "&" + u.entity + w, u.entity = "", u.state = se);
            continue;
          default:
            throw new Error(u, "Unknown state: " + u.state);
        }
      return u.position >= u.bufferCheckPosition && i(u), u;
    }
    /*! http://mths.be/fromcodepoint v0.1.0 by @mathias */
    String.fromCodePoint || function() {
      var d = String.fromCharCode, u = Math.floor, S = function() {
        var w = 16384, z = [], ne, se, Te = -1, Ae = arguments.length;
        if (!Ae)
          return "";
        for (var Ge = ""; ++Te < Ae; ) {
          var fe = Number(arguments[Te]);
          if (!isFinite(fe) || // `NaN`, `+Infinity`, or `-Infinity`
          fe < 0 || // not a valid Unicode code point
          fe > 1114111 || // not a valid Unicode code point
          u(fe) !== fe)
            throw RangeError("Invalid code point: " + fe);
          fe <= 65535 ? z.push(fe) : (fe -= 65536, ne = (fe >> 10) + 55296, se = fe % 1024 + 56320, z.push(ne, se)), (Te + 1 === Ae || z.length > w) && (Ge += d.apply(null, z), z.length = 0);
        }
        return Ge;
      };
      Object.defineProperty ? Object.defineProperty(String, "fromCodePoint", {
        value: S,
        configurable: !0,
        writable: !0
      }) : String.fromCodePoint = S;
    }();
  })(e);
})(Gc);
Object.defineProperty(ir, "__esModule", { value: !0 });
ir.XElement = void 0;
ir.parseXml = Om;
const Nm = Gc, br = gn;
class qc {
  constructor(t) {
    if (this.name = t, this.value = "", this.attributes = null, this.isCData = !1, this.elements = null, !t)
      throw (0, br.newError)("Element name cannot be empty", "ERR_XML_ELEMENT_NAME_EMPTY");
    if (!$m(t))
      throw (0, br.newError)(`Invalid element name: ${t}`, "ERR_XML_ELEMENT_INVALID_NAME");
  }
  attribute(t) {
    const n = this.attributes === null ? null : this.attributes[t];
    if (n == null)
      throw (0, br.newError)(`No attribute "${t}"`, "ERR_XML_MISSED_ATTRIBUTE");
    return n;
  }
  removeAttribute(t) {
    this.attributes !== null && delete this.attributes[t];
  }
  element(t, n = !1, r = null) {
    const i = this.elementOrNull(t, n);
    if (i === null)
      throw (0, br.newError)(r || `No element "${t}"`, "ERR_XML_MISSED_ELEMENT");
    return i;
  }
  elementOrNull(t, n = !1) {
    if (this.elements === null)
      return null;
    for (const r of this.elements)
      if (ds(r, t, n))
        return r;
    return null;
  }
  getElements(t, n = !1) {
    return this.elements === null ? [] : this.elements.filter((r) => ds(r, t, n));
  }
  elementValueOrEmpty(t, n = !1) {
    const r = this.elementOrNull(t, n);
    return r === null ? "" : r.value;
  }
}
ir.XElement = qc;
const Rm = new RegExp(/^[A-Za-z_][:A-Za-z0-9_-]*$/i);
function $m(e) {
  return Rm.test(e);
}
function ds(e, t, n) {
  const r = e.name;
  return r === t || n === !0 && r.length === t.length && r.toLowerCase() === t.toLowerCase();
}
function Om(e) {
  let t = null;
  const n = Nm.parser(!0, {}), r = [];
  return n.onopentag = (i) => {
    const o = new qc(i.name);
    if (o.attributes = i.attributes, t === null)
      t = o;
    else {
      const a = r[r.length - 1];
      a.elements == null && (a.elements = []), a.elements.push(o);
    }
    r.push(o);
  }, n.onclosetag = () => {
    r.pop();
  }, n.ontext = (i) => {
    r.length > 0 && (r[r.length - 1].value = i);
  }, n.oncdata = (i) => {
    const o = r[r.length - 1];
    o.value = i, o.isCData = !0;
  }, n.onerror = (i) => {
    throw i;
  }, n.write(e), t;
}
var Vt = {};
Object.defineProperty(Vt, "__esModule", { value: !0 });
Vt.mapToObject = Wc;
Vt.isValidKey = si;
Vt.asArray = Im;
Vt.deepAssign = Pm;
Vt.objectToArgs = xm;
function Wc(e) {
  const t = {};
  for (const [n, r] of e)
    si(n) && (r instanceof Map ? t[n] = Wc(r) : t[n] = r);
  return t;
}
function si(e) {
  return ["__proto__", "prototype", "constructor"].includes(e) ? !1 : ["string", "number", "symbol", "boolean"].includes(typeof e) || e === null;
}
function Im(e) {
  return e == null ? [] : Array.isArray(e) ? e : [e];
}
function hs(e) {
  if (Array.isArray(e))
    return !1;
  const t = typeof e;
  return t === "object" || t === "function";
}
function Dm(e, t, n) {
  const r = t[n];
  if (r === void 0)
    return;
  const i = e[n];
  i == null || r == null || !hs(i) || !hs(r) ? Array.isArray(i) && Array.isArray(r) ? e[n] = Array.from(new Set(i.concat(r))) : e[n] = r : e[n] = Vc(i, r);
}
function Vc(e, t) {
  if (e !== t)
    for (const n of Object.getOwnPropertyNames(t))
      si(n) && Dm(e, t, n);
  return e;
}
function Pm(e, ...t) {
  for (const n of t)
    n != null && Vc(e, n);
  return e;
}
const Lm = /^[a-zA-Z][a-zA-Z0-9-]*$/, Fm = /[\0\r\n]/;
function xm(e) {
  const t = Object.entries(e).reduce((n, [r, i]) => {
    if (!si(r) || i == null)
      return n;
    if (!Lm.test(r))
      throw new Error(`objectToArgs: unsafe flag name rejected: ${JSON.stringify(r)}`);
    if (Fm.test(i))
      throw new Error(`objectToArgs: value for --${r} contains a null byte or newline`);
    return n.concat([`--${r}`, i]);
  }, []);
  return Object.freeze(t);
}
(function(e) {
  Object.defineProperty(e, "__esModule", { value: !0 }), e.CURRENT_APP_PACKAGE_FILE_NAME = e.CURRENT_APP_INSTALLER_FILE_NAME = e.objectToArgs = e.deepAssign = e.asArray = e.mapToObject = e.isValidKey = e.XElement = e.parseXml = e.UUID = e.parseDn = e.retry = e.githubTagPrefix = e.githubUrl = e.getS3LikeProviderBaseUrl = e.ProgressCallbackTransform = e.MemoLazy = e.safeStringifyJson = e.safeGetHeader = e.parseJson = e.isSensitiveFieldName = e.HttpExecutor = e.hashSensitiveValue = e.HttpError = e.DigestTransform = e.createHttpError = e.configureRequestUrl = e.configureRequestOptionsFromUrl = e.configureRequestOptions = e.newError = e.CancellationToken = e.CancellationError = void 0;
  var t = Tt;
  Object.defineProperty(e, "CancellationError", { enumerable: !0, get: function() {
    return t.CancellationError;
  } }), Object.defineProperty(e, "CancellationToken", { enumerable: !0, get: function() {
    return t.CancellationToken;
  } });
  var n = gn;
  Object.defineProperty(e, "newError", { enumerable: !0, get: function() {
    return n.newError;
  } });
  var r = de;
  Object.defineProperty(e, "configureRequestOptions", { enumerable: !0, get: function() {
    return r.configureRequestOptions;
  } }), Object.defineProperty(e, "configureRequestOptionsFromUrl", { enumerable: !0, get: function() {
    return r.configureRequestOptionsFromUrl;
  } }), Object.defineProperty(e, "configureRequestUrl", { enumerable: !0, get: function() {
    return r.configureRequestUrl;
  } }), Object.defineProperty(e, "createHttpError", { enumerable: !0, get: function() {
    return r.createHttpError;
  } }), Object.defineProperty(e, "DigestTransform", { enumerable: !0, get: function() {
    return r.DigestTransform;
  } }), Object.defineProperty(e, "HttpError", { enumerable: !0, get: function() {
    return r.HttpError;
  } }), Object.defineProperty(e, "hashSensitiveValue", { enumerable: !0, get: function() {
    return r.hashSensitiveValue;
  } }), Object.defineProperty(e, "HttpExecutor", { enumerable: !0, get: function() {
    return r.HttpExecutor;
  } }), Object.defineProperty(e, "isSensitiveFieldName", { enumerable: !0, get: function() {
    return r.isSensitiveFieldName;
  } }), Object.defineProperty(e, "parseJson", { enumerable: !0, get: function() {
    return r.parseJson;
  } }), Object.defineProperty(e, "safeGetHeader", { enumerable: !0, get: function() {
    return r.safeGetHeader;
  } }), Object.defineProperty(e, "safeStringifyJson", { enumerable: !0, get: function() {
    return r.safeStringifyJson;
  } });
  var i = ai;
  Object.defineProperty(e, "MemoLazy", { enumerable: !0, get: function() {
    return i.MemoLazy;
  } });
  var o = nr;
  Object.defineProperty(e, "ProgressCallbackTransform", { enumerable: !0, get: function() {
    return o.ProgressCallbackTransform;
  } });
  var a = rr;
  Object.defineProperty(e, "getS3LikeProviderBaseUrl", { enumerable: !0, get: function() {
    return a.getS3LikeProviderBaseUrl;
  } }), Object.defineProperty(e, "githubUrl", { enumerable: !0, get: function() {
    return a.githubUrl;
  } }), Object.defineProperty(e, "githubTagPrefix", { enumerable: !0, get: function() {
    return a.githubTagPrefix;
  } });
  var s = Vo;
  Object.defineProperty(e, "retry", { enumerable: !0, get: function() {
    return s.retry;
  } });
  var c = zo;
  Object.defineProperty(e, "parseDn", { enumerable: !0, get: function() {
    return c.parseDn;
  } });
  var p = hn;
  Object.defineProperty(e, "UUID", { enumerable: !0, get: function() {
    return p.UUID;
  } });
  var l = ir;
  Object.defineProperty(e, "parseXml", { enumerable: !0, get: function() {
    return l.parseXml;
  } }), Object.defineProperty(e, "XElement", { enumerable: !0, get: function() {
    return l.XElement;
  } });
  var f = Vt;
  Object.defineProperty(e, "isValidKey", { enumerable: !0, get: function() {
    return f.isValidKey;
  } }), Object.defineProperty(e, "mapToObject", { enumerable: !0, get: function() {
    return f.mapToObject;
  } }), Object.defineProperty(e, "asArray", { enumerable: !0, get: function() {
    return f.asArray;
  } }), Object.defineProperty(e, "deepAssign", { enumerable: !0, get: function() {
    return f.deepAssign;
  } }), Object.defineProperty(e, "objectToArgs", { enumerable: !0, get: function() {
    return f.objectToArgs;
  } }), e.CURRENT_APP_INSTALLER_FILE_NAME = "installer.exe", e.CURRENT_APP_PACKAGE_FILE_NAME = "package.7z";
})(he);
var we = {}, Yo = {}, Ke = {};
function zc(e) {
  return typeof e > "u" || e === null;
}
function Um(e) {
  return typeof e == "object" && e !== null;
}
function km(e) {
  return Array.isArray(e) ? e : zc(e) ? [] : [e];
}
function Mm(e, t) {
  if (t) {
    const n = Object.keys(t);
    for (let r = 0, i = n.length; r < i; r += 1) {
      const o = n[r];
      e[o] = t[o];
    }
  }
  return e;
}
function Bm(e, t) {
  let n = "";
  for (let r = 0; r < t; r += 1)
    n += e;
  return n;
}
function Hm(e) {
  return e === 0 && Number.NEGATIVE_INFINITY === 1 / e;
}
Ke.isNothing = zc;
Ke.isObject = Um;
Ke.toArray = km;
Ke.repeat = Bm;
Ke.isNegativeZero = Hm;
Ke.extend = Mm;
function Yc(e, t) {
  let n = "";
  const r = e.reason || "(unknown reason)";
  return e.mark ? (e.mark.name && (n += 'in "' + e.mark.name + '" '), n += "(" + (e.mark.line + 1) + ":" + (e.mark.column + 1) + ")", !t && e.mark.snippet && (n += `

` + e.mark.snippet), r + " " + n) : r;
}
function Gn(e, t) {
  Error.call(this), this.name = "YAMLException", this.reason = e, this.mark = t, this.message = Yc(this, !1), Error.captureStackTrace ? Error.captureStackTrace(this, this.constructor) : this.stack = new Error().stack || "";
}
Gn.prototype = Object.create(Error.prototype);
Gn.prototype.constructor = Gn;
Gn.prototype.toString = function(t) {
  return this.name + ": " + Yc(this, t);
};
var or = Gn;
const On = Ke;
function ji(e, t, n, r, i) {
  let o = "", a = "";
  const s = Math.floor(i / 2) - 1;
  return r - t > s && (o = " ... ", t = r - s + o.length), n - r > s && (a = " ...", n = r + s - a.length), {
    str: o + e.slice(t, n).replace(/\t/g, "→") + a,
    pos: r - t + o.length
    // relative position
  };
}
function Gi(e, t) {
  return On.repeat(" ", t - e.length) + e;
}
function jm(e, t) {
  if (t = Object.create(t || null), !e.buffer) return null;
  t.maxLength || (t.maxLength = 79), typeof t.indent != "number" && (t.indent = 1), typeof t.linesBefore != "number" && (t.linesBefore = 3), typeof t.linesAfter != "number" && (t.linesAfter = 2);
  const n = /\r?\n|\r|\0/g, r = [0], i = [];
  let o, a = -1;
  for (; o = n.exec(e.buffer); )
    i.push(o.index), r.push(o.index + o[0].length), e.position <= o.index && a < 0 && (a = r.length - 2);
  a < 0 && (a = r.length - 1);
  let s = "";
  const c = Math.min(e.line + t.linesAfter, i.length).toString().length, p = t.maxLength - (t.indent + c + 3);
  for (let f = 1; f <= t.linesBefore && !(a - f < 0); f++) {
    const h = ji(
      e.buffer,
      r[a - f],
      i[a - f],
      e.position - (r[a] - r[a - f]),
      p
    );
    s = On.repeat(" ", t.indent) + Gi((e.line - f + 1).toString(), c) + " | " + h.str + `
` + s;
  }
  const l = ji(e.buffer, r[a], i[a], e.position, p);
  s += On.repeat(" ", t.indent) + Gi((e.line + 1).toString(), c) + " | " + l.str + `
`, s += On.repeat("-", t.indent + c + 3 + l.pos) + `^
`;
  for (let f = 1; f <= t.linesAfter && !(a + f >= i.length); f++) {
    const h = ji(
      e.buffer,
      r[a + f],
      i[a + f],
      e.position - (r[a] - r[a + f]),
      p
    );
    s += On.repeat(" ", t.indent) + Gi((e.line + f + 1).toString(), c) + " | " + h.str + `
`;
  }
  return s.replace(/\n$/, "");
}
var Gm = jm;
const ps = or, qm = [
  "kind",
  "multi",
  "resolve",
  "construct",
  "instanceOf",
  "predicate",
  "represent",
  "representName",
  "defaultStyle",
  "styleAliases"
], Wm = [
  "scalar",
  "sequence",
  "mapping"
];
function Vm(e) {
  const t = {};
  return e !== null && Object.keys(e).forEach(function(n) {
    e[n].forEach(function(r) {
      t[String(r)] = n;
    });
  }), t;
}
function zm(e, t) {
  if (t = t || {}, Object.keys(t).forEach(function(n) {
    if (qm.indexOf(n) === -1)
      throw new ps('Unknown option "' + n + '" is met in definition of "' + e + '" YAML type.');
  }), this.options = t, this.tag = e, this.kind = t.kind || null, this.resolve = t.resolve || function() {
    return !0;
  }, this.construct = t.construct || function(n) {
    return n;
  }, this.instanceOf = t.instanceOf || null, this.predicate = t.predicate || null, this.represent = t.represent || null, this.representName = t.representName || null, this.defaultStyle = t.defaultStyle || null, this.multi = t.multi || !1, this.styleAliases = Vm(t.styleAliases || null), Wm.indexOf(this.kind) === -1)
    throw new ps('Unknown kind "' + this.kind + '" is specified for "' + e + '" YAML type.');
}
var Le = zm;
const Nn = or, qi = Le;
function ms(e, t) {
  const n = [];
  return e[t].forEach(function(r) {
    let i = n.length;
    n.forEach(function(o, a) {
      o.tag === r.tag && o.kind === r.kind && o.multi === r.multi && (i = a);
    }), n[i] = r;
  }), n;
}
function Ym() {
  const e = {
    scalar: {},
    sequence: {},
    mapping: {},
    fallback: {},
    multi: {
      scalar: [],
      sequence: [],
      mapping: [],
      fallback: []
    }
  };
  function t(n) {
    n.multi ? (e.multi[n.kind].push(n), e.multi.fallback.push(n)) : e[n.kind][n.tag] = e.fallback[n.tag] = n;
  }
  for (let n = 0, r = arguments.length; n < r; n += 1)
    arguments[n].forEach(t);
  return e;
}
function So(e) {
  return this.extend(e);
}
So.prototype.extend = function(t) {
  let n = [], r = [];
  if (t instanceof qi)
    r.push(t);
  else if (Array.isArray(t))
    r = r.concat(t);
  else if (t && (Array.isArray(t.implicit) || Array.isArray(t.explicit)))
    t.implicit && (n = n.concat(t.implicit)), t.explicit && (r = r.concat(t.explicit));
  else
    throw new Nn("Schema.extend argument should be a Type, [ Type ], or a schema definition ({ implicit: [...], explicit: [...] })");
  n.forEach(function(o) {
    if (!(o instanceof qi))
      throw new Nn("Specified list of YAML types (or a single Type object) contains a non-Type object.");
    if (o.loadKind && o.loadKind !== "scalar")
      throw new Nn("There is a non-scalar type in the implicit list of a schema. Implicit resolving of such types is not supported.");
    if (o.multi)
      throw new Nn("There is a multi type in the implicit list of a schema. Multi tags can only be listed as explicit.");
  }), r.forEach(function(o) {
    if (!(o instanceof qi))
      throw new Nn("Specified list of YAML types (or a single Type object) contains a non-Type object.");
  });
  const i = Object.create(So.prototype);
  return i.implicit = (this.implicit || []).concat(n), i.explicit = (this.explicit || []).concat(r), i.compiledImplicit = ms(i, "implicit"), i.compiledExplicit = ms(i, "explicit"), i.compiledTypeMap = Ym(i.compiledImplicit, i.compiledExplicit), i;
};
var Xc = So;
const Xm = Le;
var Kc = new Xm("tag:yaml.org,2002:str", {
  kind: "scalar",
  construct: function(e) {
    return e !== null ? e : "";
  }
});
const Km = Le;
var Jc = new Km("tag:yaml.org,2002:seq", {
  kind: "sequence",
  construct: function(e) {
    return e !== null ? e : [];
  }
});
const Jm = Le;
var Qc = new Jm("tag:yaml.org,2002:map", {
  kind: "mapping",
  construct: function(e) {
    return e !== null ? e : {};
  }
});
const Qm = Xc;
var Zc = new Qm({
  explicit: [
    Kc,
    Jc,
    Qc
  ]
});
const Zm = Le;
function eg(e) {
  if (e === null) return !0;
  const t = e.length;
  return t === 1 && e === "~" || t === 4 && (e === "null" || e === "Null" || e === "NULL");
}
function tg() {
  return null;
}
function ng(e) {
  return e === null;
}
var eu = new Zm("tag:yaml.org,2002:null", {
  kind: "scalar",
  resolve: eg,
  construct: tg,
  predicate: ng,
  represent: {
    canonical: function() {
      return "~";
    },
    lowercase: function() {
      return "null";
    },
    uppercase: function() {
      return "NULL";
    },
    camelcase: function() {
      return "Null";
    },
    empty: function() {
      return "";
    }
  },
  defaultStyle: "lowercase"
});
const rg = Le;
function ig(e) {
  if (e === null) return !1;
  const t = e.length;
  return t === 4 && (e === "true" || e === "True" || e === "TRUE") || t === 5 && (e === "false" || e === "False" || e === "FALSE");
}
function og(e) {
  return e === "true" || e === "True" || e === "TRUE";
}
function ag(e) {
  return Object.prototype.toString.call(e) === "[object Boolean]";
}
var tu = new rg("tag:yaml.org,2002:bool", {
  kind: "scalar",
  resolve: ig,
  construct: og,
  predicate: ag,
  represent: {
    lowercase: function(e) {
      return e ? "true" : "false";
    },
    uppercase: function(e) {
      return e ? "TRUE" : "FALSE";
    },
    camelcase: function(e) {
      return e ? "True" : "False";
    }
  },
  defaultStyle: "lowercase"
});
const sg = Ke, lg = Le;
function cg(e) {
  return e >= 48 && e <= 57 || e >= 65 && e <= 70 || e >= 97 && e <= 102;
}
function ug(e) {
  return e >= 48 && e <= 55;
}
function fg(e) {
  return e >= 48 && e <= 57;
}
function dg(e) {
  if (e === null) return !1;
  const t = e.length;
  let n = 0, r = !1;
  if (!t) return !1;
  let i = e[n];
  if ((i === "-" || i === "+") && (i = e[++n]), i === "0") {
    if (n + 1 === t) return !0;
    if (i = e[++n], i === "b") {
      for (n++; n < t; n++) {
        if (i = e[n], i !== "0" && i !== "1") return !1;
        r = !0;
      }
      return r && isFinite(In(e));
    }
    if (i === "x") {
      for (n++; n < t; n++) {
        if (!cg(e.charCodeAt(n))) return !1;
        r = !0;
      }
      return r && isFinite(In(e));
    }
    if (i === "o") {
      for (n++; n < t; n++) {
        if (!ug(e.charCodeAt(n))) return !1;
        r = !0;
      }
      return r && isFinite(In(e));
    }
  }
  for (; n < t; n++) {
    if (!fg(e.charCodeAt(n)))
      return !1;
    r = !0;
  }
  return r ? isFinite(In(e)) : !1;
}
function In(e) {
  let t = e, n = 1, r = t[0];
  if ((r === "-" || r === "+") && (r === "-" && (n = -1), t = t.slice(1), r = t[0]), t === "0") return 0;
  if (r === "0") {
    if (t[1] === "b") return n * parseInt(t.slice(2), 2);
    if (t[1] === "x") return n * parseInt(t.slice(2), 16);
    if (t[1] === "o") return n * parseInt(t.slice(2), 8);
  }
  return n * parseInt(t, 10);
}
function hg(e) {
  return In(e);
}
function pg(e) {
  return Object.prototype.toString.call(e) === "[object Number]" && e % 1 === 0 && !sg.isNegativeZero(e);
}
var nu = new lg("tag:yaml.org,2002:int", {
  kind: "scalar",
  resolve: dg,
  construct: hg,
  predicate: pg,
  represent: {
    binary: function(e) {
      return e >= 0 ? "0b" + e.toString(2) : "-0b" + e.toString(2).slice(1);
    },
    octal: function(e) {
      return e >= 0 ? "0o" + e.toString(8) : "-0o" + e.toString(8).slice(1);
    },
    decimal: function(e) {
      return e.toString(10);
    },
    hexadecimal: function(e) {
      return e >= 0 ? "0x" + e.toString(16).toUpperCase() : "-0x" + e.toString(16).toUpperCase().slice(1);
    }
  },
  defaultStyle: "decimal",
  styleAliases: {
    binary: [2, "bin"],
    octal: [8, "oct"],
    decimal: [10, "dec"],
    hexadecimal: [16, "hex"]
  }
});
const ru = Ke, mg = Le, gg = new RegExp(
  // 2.5e4, 2.5 and integers
  "^(?:[-+]?(?:[0-9]+)(?:\\.[0-9]*)?(?:[eE][-+]?[0-9]+)?|\\.[0-9]+(?:[eE][-+]?[0-9]+)?|[-+]?\\.(?:inf|Inf|INF)|\\.(?:nan|NaN|NAN))$"
), Eg = new RegExp(
  "^(?:[-+]?\\.(?:inf|Inf|INF)|\\.(?:nan|NaN|NAN))$"
);
function yg(e) {
  return e === null || !gg.test(e) ? !1 : isFinite(parseFloat(e, 10)) ? !0 : Eg.test(e);
}
function _g(e) {
  let t = e.toLowerCase();
  const n = t[0] === "-" ? -1 : 1;
  return "+-".indexOf(t[0]) >= 0 && (t = t.slice(1)), t === ".inf" ? n === 1 ? Number.POSITIVE_INFINITY : Number.NEGATIVE_INFINITY : t === ".nan" ? NaN : n * parseFloat(t, 10);
}
const wg = /^[-+]?[0-9]+e/;
function vg(e, t) {
  if (isNaN(e))
    switch (t) {
      case "lowercase":
        return ".nan";
      case "uppercase":
        return ".NAN";
      case "camelcase":
        return ".NaN";
    }
  else if (Number.POSITIVE_INFINITY === e)
    switch (t) {
      case "lowercase":
        return ".inf";
      case "uppercase":
        return ".INF";
      case "camelcase":
        return ".Inf";
    }
  else if (Number.NEGATIVE_INFINITY === e)
    switch (t) {
      case "lowercase":
        return "-.inf";
      case "uppercase":
        return "-.INF";
      case "camelcase":
        return "-.Inf";
    }
  else if (ru.isNegativeZero(e))
    return "-0.0";
  const n = e.toString(10);
  return wg.test(n) ? n.replace("e", ".e") : n;
}
function Tg(e) {
  return Object.prototype.toString.call(e) === "[object Number]" && (e % 1 !== 0 || ru.isNegativeZero(e));
}
var iu = new mg("tag:yaml.org,2002:float", {
  kind: "scalar",
  resolve: yg,
  construct: _g,
  predicate: Tg,
  represent: vg,
  defaultStyle: "lowercase"
}), ou = Zc.extend({
  implicit: [
    eu,
    tu,
    nu,
    iu
  ]
}), au = ou;
const Ag = Le, su = new RegExp(
  "^([0-9][0-9][0-9][0-9])-([0-9][0-9])-([0-9][0-9])$"
), lu = new RegExp(
  "^([0-9][0-9][0-9][0-9])-([0-9][0-9]?)-([0-9][0-9]?)(?:[Tt]|[ \\t]+)([0-9][0-9]?):([0-9][0-9]):([0-9][0-9])(?:\\.([0-9]*))?(?:[ \\t]*(Z|([-+])([0-9][0-9]?)(?::([0-9][0-9]))?))?$"
);
function Sg(e) {
  return e === null ? !1 : su.exec(e) !== null || lu.exec(e) !== null;
}
function Cg(e) {
  let t = 0, n = null, r = su.exec(e);
  if (r === null && (r = lu.exec(e)), r === null) throw new Error("Date resolve error");
  const i = +r[1], o = +r[2] - 1, a = +r[3];
  if (!r[4])
    return new Date(Date.UTC(i, o, a));
  const s = +r[4], c = +r[5], p = +r[6];
  if (r[7]) {
    for (t = r[7].slice(0, 3); t.length < 3; )
      t += "0";
    t = +t;
  }
  if (r[9]) {
    const f = +r[10], h = +(r[11] || 0);
    n = (f * 60 + h) * 6e4, r[9] === "-" && (n = -n);
  }
  const l = new Date(Date.UTC(i, o, a, s, c, p, t));
  return n && l.setTime(l.getTime() - n), l;
}
function bg(e) {
  return e.toISOString();
}
var cu = new Ag("tag:yaml.org,2002:timestamp", {
  kind: "scalar",
  resolve: Sg,
  construct: Cg,
  instanceOf: Date,
  represent: bg
});
const Ng = Le;
function Rg(e) {
  return e === "<<" || e === null;
}
var uu = new Ng("tag:yaml.org,2002:merge", {
  kind: "scalar",
  resolve: Rg
});
const $g = Le, Xo = `ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/=
\r`;
function Og(e) {
  if (e === null) return !1;
  let t = 0;
  const n = e.length, r = Xo;
  for (let i = 0; i < n; i++) {
    const o = r.indexOf(e.charAt(i));
    if (!(o > 64)) {
      if (o < 0) return !1;
      t += 6;
    }
  }
  return t % 8 === 0;
}
function Ig(e) {
  const t = e.replace(/[\r\n=]/g, ""), n = t.length, r = Xo;
  let i = 0;
  const o = [];
  for (let s = 0; s < n; s++)
    s % 4 === 0 && s && (o.push(i >> 16 & 255), o.push(i >> 8 & 255), o.push(i & 255)), i = i << 6 | r.indexOf(t.charAt(s));
  const a = n % 4 * 6;
  return a === 0 ? (o.push(i >> 16 & 255), o.push(i >> 8 & 255), o.push(i & 255)) : a === 18 ? (o.push(i >> 10 & 255), o.push(i >> 2 & 255)) : a === 12 && o.push(i >> 4 & 255), new Uint8Array(o);
}
function Dg(e) {
  let t = "", n = 0;
  const r = e.length, i = Xo;
  for (let a = 0; a < r; a++)
    a % 3 === 0 && a && (t += i[n >> 18 & 63], t += i[n >> 12 & 63], t += i[n >> 6 & 63], t += i[n & 63]), n = (n << 8) + e[a];
  const o = r % 3;
  return o === 0 ? (t += i[n >> 18 & 63], t += i[n >> 12 & 63], t += i[n >> 6 & 63], t += i[n & 63]) : o === 2 ? (t += i[n >> 10 & 63], t += i[n >> 4 & 63], t += i[n << 2 & 63], t += i[64]) : o === 1 && (t += i[n >> 2 & 63], t += i[n << 4 & 63], t += i[64], t += i[64]), t;
}
function Pg(e) {
  return Object.prototype.toString.call(e) === "[object Uint8Array]";
}
var fu = new $g("tag:yaml.org,2002:binary", {
  kind: "scalar",
  resolve: Og,
  construct: Ig,
  predicate: Pg,
  represent: Dg
});
const Lg = Le, Fg = Object.prototype.hasOwnProperty, xg = Object.prototype.toString;
function Ug(e) {
  if (e === null) return !0;
  const t = [], n = e;
  for (let r = 0, i = n.length; r < i; r += 1) {
    const o = n[r];
    let a = !1;
    if (xg.call(o) !== "[object Object]") return !1;
    let s;
    for (s in o)
      if (Fg.call(o, s))
        if (!a) a = !0;
        else return !1;
    if (!a) return !1;
    if (t.indexOf(s) === -1) t.push(s);
    else return !1;
  }
  return !0;
}
function kg(e) {
  return e !== null ? e : [];
}
var du = new Lg("tag:yaml.org,2002:omap", {
  kind: "sequence",
  resolve: Ug,
  construct: kg
});
const Mg = Le, Bg = Object.prototype.toString;
function Hg(e) {
  if (e === null) return !0;
  const t = e, n = new Array(t.length);
  for (let r = 0, i = t.length; r < i; r += 1) {
    const o = t[r];
    if (Bg.call(o) !== "[object Object]") return !1;
    const a = Object.keys(o);
    if (a.length !== 1) return !1;
    n[r] = [a[0], o[a[0]]];
  }
  return !0;
}
function jg(e) {
  if (e === null) return [];
  const t = e, n = new Array(t.length);
  for (let r = 0, i = t.length; r < i; r += 1) {
    const o = t[r], a = Object.keys(o);
    n[r] = [a[0], o[a[0]]];
  }
  return n;
}
var hu = new Mg("tag:yaml.org,2002:pairs", {
  kind: "sequence",
  resolve: Hg,
  construct: jg
});
const Gg = Le, qg = Object.prototype.hasOwnProperty;
function Wg(e) {
  if (e === null) return !0;
  const t = e;
  for (const n in t)
    if (qg.call(t, n) && t[n] !== null)
      return !1;
  return !0;
}
function Vg(e) {
  return e !== null ? e : {};
}
var pu = new Gg("tag:yaml.org,2002:set", {
  kind: "mapping",
  resolve: Wg,
  construct: Vg
}), Ko = au.extend({
  implicit: [
    cu,
    uu
  ],
  explicit: [
    fu,
    du,
    hu,
    pu
  ]
});
const Ut = Ke, mu = or, zg = Gm, Yg = Ko, Xe = Object.prototype.hasOwnProperty, Vr = 1, gu = 2, Eu = 3, zr = 4, Wi = 1, Xg = 2, gs = 3, Kg = /[\x00-\x08\x0B\x0C\x0E-\x1F\x7F-\x84\x86-\x9F\uFFFE\uFFFF]|[\uD800-\uDBFF](?![\uDC00-\uDFFF])|(?:[^\uD800-\uDBFF]|^)[\uDC00-\uDFFF]/, Jg = /[\x85\u2028\u2029]/, Qg = /[,\[\]{}]/, yu = /^(?:!|!!|![0-9A-Za-z-]+!)$/, _u = /^(?:!|[^,\[\]{}])(?:%[0-9a-f]{2}|[0-9a-z\-#;/?:@&=+$,_.!~*'()\[\]])*$/i;
function Es(e) {
  return Object.prototype.toString.call(e);
}
function rt(e) {
  return e === 10 || e === 13;
}
function ct(e) {
  return e === 9 || e === 32;
}
function Ue(e) {
  return e === 9 || e === 32 || e === 10 || e === 13;
}
function on(e) {
  return e === 44 || e === 91 || e === 93 || e === 123 || e === 125;
}
function Zg(e) {
  if (e >= 48 && e <= 57)
    return e - 48;
  const t = e | 32;
  return t >= 97 && t <= 102 ? t - 97 + 10 : -1;
}
function e0(e) {
  return e === 120 ? 2 : e === 117 ? 4 : e === 85 ? 8 : 0;
}
function t0(e) {
  return e >= 48 && e <= 57 ? e - 48 : -1;
}
function ys(e) {
  switch (e) {
    case 48:
      return "\0";
    case 97:
      return "\x07";
    case 98:
      return "\b";
    case 116:
      return "	";
    case 9:
      return "	";
    case 110:
      return `
`;
    case 118:
      return "\v";
    case 102:
      return "\f";
    case 114:
      return "\r";
    case 101:
      return "\x1B";
    case 32:
      return " ";
    case 34:
      return '"';
    case 47:
      return "/";
    case 92:
      return "\\";
    case 78:
      return "";
    case 95:
      return " ";
    case 76:
      return "\u2028";
    case 80:
      return "\u2029";
    default:
      return "";
  }
}
function n0(e) {
  return e <= 65535 ? String.fromCharCode(e) : String.fromCharCode(
    (e - 65536 >> 10) + 55296,
    (e - 65536 & 1023) + 56320
  );
}
function wu(e, t, n) {
  t === "__proto__" ? Object.defineProperty(e, t, {
    configurable: !0,
    enumerable: !0,
    writable: !0,
    value: n
  }) : e[t] = n;
}
const vu = new Array(256), Tu = new Array(256);
for (let e = 0; e < 256; e++)
  vu[e] = ys(e) ? 1 : 0, Tu[e] = ys(e);
function r0(e, t) {
  this.input = e, this.filename = t.filename || null, this.schema = t.schema || Yg, this.onWarning = t.onWarning || null, this.legacy = t.legacy || !1, this.json = t.json || !1, this.listener = t.listener || null, this.maxDepth = typeof t.maxDepth == "number" ? t.maxDepth : 100, this.maxTotalMergeKeys = typeof t.maxTotalMergeKeys == "number" ? t.maxTotalMergeKeys : 1e4, this.implicitTypes = this.schema.compiledImplicit, this.typeMap = this.schema.compiledTypeMap, this.length = e.length, this.position = 0, this.line = 0, this.lineStart = 0, this.lineIndent = 0, this.depth = 0, this.totalMergeKeys = 0, this.firstTabInLine = -1, this.documents = [], this.anchorMapTransactions = [];
}
function Au(e, t) {
  const n = {
    name: e.filename,
    buffer: e.input.slice(0, -1),
    // omit trailing \0
    position: e.position,
    line: e.line,
    column: e.position - e.lineStart
  };
  return n.snippet = zg(n), new mu(t, n);
}
function x(e, t) {
  throw Au(e, t);
}
function Yr(e, t) {
  e.onWarning && e.onWarning.call(null, Au(e, t));
}
function kt(e, t, n) {
  const r = e.anchorMapTransactions;
  if (r.length !== 0) {
    const i = r[r.length - 1];
    Xe.call(i, t) || (i[t] = {
      existed: Xe.call(e.anchorMap, t),
      value: e.anchorMap[t]
    });
  }
  e.anchorMap[t] = n;
}
function i0(e) {
  e.anchorMapTransactions.push(/* @__PURE__ */ Object.create(null));
}
function o0(e) {
  const t = e.anchorMapTransactions.pop(), n = e.anchorMapTransactions;
  if (n.length === 0) return;
  const r = n[n.length - 1], i = Object.keys(t);
  for (let o = 0, a = i.length; o < a; o += 1) {
    const s = i[o];
    Xe.call(r, s) || (r[s] = t[s]);
  }
}
function a0(e) {
  const t = e.anchorMapTransactions.pop(), n = Object.keys(t);
  for (let r = n.length - 1; r >= 0; r -= 1) {
    const i = t[n[r]];
    i.existed ? e.anchorMap[n[r]] = i.value : delete e.anchorMap[n[r]];
  }
}
function Su(e) {
  return {
    position: e.position,
    line: e.line,
    lineStart: e.lineStart,
    lineIndent: e.lineIndent,
    firstTabInLine: e.firstTabInLine,
    tag: e.tag,
    anchor: e.anchor,
    kind: e.kind,
    result: e.result
  };
}
function _s(e, t) {
  e.position = t.position, e.line = t.line, e.lineStart = t.lineStart, e.lineIndent = t.lineIndent, e.firstTabInLine = t.firstTabInLine, e.tag = t.tag, e.anchor = t.anchor, e.kind = t.kind, e.result = t.result;
}
const ws = {
  YAML: function(t, n, r) {
    t.version !== null && x(t, "duplication of %YAML directive"), r.length !== 1 && x(t, "YAML directive accepts exactly one argument");
    const i = /^([0-9]+)\.([0-9]+)$/.exec(r[0]);
    i === null && x(t, "ill-formed argument of the YAML directive");
    const o = parseInt(i[1], 10), a = parseInt(i[2], 10);
    o !== 1 && x(t, "unacceptable YAML version of the document"), t.version = r[0], t.checkLineBreaks = a < 2, a !== 1 && a !== 2 && Yr(t, "unsupported YAML version of the document");
  },
  TAG: function(t, n, r) {
    let i;
    r.length !== 2 && x(t, "TAG directive accepts exactly two arguments");
    const o = r[0];
    i = r[1], yu.test(o) || x(t, "ill-formed tag handle (first argument) of the TAG directive"), Xe.call(t.tagMap, o) && x(t, 'there is a previously declared suffix for "' + o + '" tag handle'), _u.test(i) || x(t, "ill-formed tag prefix (second argument) of the TAG directive");
    try {
      i = decodeURIComponent(i);
    } catch {
      x(t, "tag prefix is malformed: " + i);
    }
    t.tagMap[o] = i;
  }
};
function wt(e, t, n, r) {
  if (t < n) {
    const i = e.input.slice(t, n);
    if (r)
      for (let o = 0, a = i.length; o < a; o += 1) {
        const s = i.charCodeAt(o);
        s === 9 || s >= 32 && s <= 1114111 || x(e, "expected valid JSON character");
      }
    else Kg.test(i) && x(e, "the stream contains non-printable characters");
    e.result += i;
  }
}
function vs(e, t, n, r) {
  Ut.isObject(n) || x(e, "cannot merge mappings; the provided source object is unacceptable");
  const i = Object.keys(n);
  for (let o = 0, a = i.length; o < a; o += 1) {
    const s = i[o];
    e.maxTotalMergeKeys !== -1 && ++e.totalMergeKeys > e.maxTotalMergeKeys && x(e, "merge keys exceeded maxTotalMergeKeys (" + e.maxTotalMergeKeys + ")"), Xe.call(t, s) || (wu(t, s, n[s]), r[s] = !0);
  }
}
function an(e, t, n, r, i, o, a, s, c) {
  if (Array.isArray(i)) {
    i = Array.prototype.slice.call(i);
    for (let p = 0, l = i.length; p < l; p += 1)
      Array.isArray(i[p]) && x(e, "nested arrays are not supported inside keys"), typeof i == "object" && Es(i[p]) === "[object Object]" && (i[p] = "[object Object]");
  }
  if (typeof i == "object" && Es(i) === "[object Object]" && (i = "[object Object]"), i = String(i), t === null && (t = {}), r === "tag:yaml.org,2002:merge")
    if (Array.isArray(o))
      for (let p = 0, l = o.length; p < l; p += 1)
        vs(e, t, o[p], n);
    else
      vs(e, t, o, n);
  else
    !e.json && !Xe.call(n, i) && Xe.call(t, i) && (e.line = a || e.line, e.lineStart = s || e.lineStart, e.position = c || e.position, x(e, "duplicated mapping key")), wu(t, i, o), delete n[i];
  return t;
}
function Jo(e) {
  const t = e.input.charCodeAt(e.position);
  t === 10 ? e.position++ : t === 13 ? (e.position++, e.input.charCodeAt(e.position) === 10 && e.position++) : x(e, "a line break is expected"), e.line += 1, e.lineStart = e.position, e.firstTabInLine = -1;
}
function le(e, t, n) {
  let r = 0, i = e.input.charCodeAt(e.position);
  for (; i !== 0; ) {
    for (; ct(i); )
      i === 9 && e.firstTabInLine === -1 && (e.firstTabInLine = e.position), i = e.input.charCodeAt(++e.position);
    if (t && i === 35)
      do
        i = e.input.charCodeAt(++e.position);
      while (i !== 10 && i !== 13 && i !== 0);
    if (rt(i))
      for (Jo(e), i = e.input.charCodeAt(e.position), r++, e.lineIndent = 0; i === 32; )
        e.lineIndent++, i = e.input.charCodeAt(++e.position);
    else
      break;
  }
  return n !== -1 && r !== 0 && e.lineIndent < n && Yr(e, "deficient indentation"), r;
}
function li(e) {
  let t = e.position, n = e.input.charCodeAt(t);
  return !!((n === 45 || n === 46) && n === e.input.charCodeAt(t + 1) && n === e.input.charCodeAt(t + 2) && (t += 3, n = e.input.charCodeAt(t), n === 0 || Ue(n)));
}
function Qo(e, t) {
  t === 1 ? e.result += " " : t > 1 && (e.result += Ut.repeat(`
`, t - 1));
}
function s0(e, t, n) {
  let r, i, o, a, s, c;
  const p = e.kind, l = e.result;
  let f = e.input.charCodeAt(e.position);
  if (Ue(f) || on(f) || f === 35 || f === 38 || f === 42 || f === 33 || f === 124 || f === 62 || f === 39 || f === 34 || f === 37 || f === 64 || f === 96)
    return !1;
  if (f === 63 || f === 45) {
    const h = e.input.charCodeAt(e.position + 1);
    if (Ue(h) || n && on(h))
      return !1;
  }
  for (e.kind = "scalar", e.result = "", r = i = e.position, o = !1; f !== 0; ) {
    if (f === 58) {
      const h = e.input.charCodeAt(e.position + 1);
      if (Ue(h) || n && on(h))
        break;
    } else if (f === 35) {
      const h = e.input.charCodeAt(e.position - 1);
      if (Ue(h))
        break;
    } else {
      if (e.position === e.lineStart && li(e) || n && on(f))
        break;
      if (rt(f))
        if (a = e.line, s = e.lineStart, c = e.lineIndent, le(e, !1, -1), e.lineIndent >= t) {
          o = !0, f = e.input.charCodeAt(e.position);
          continue;
        } else {
          e.position = i, e.line = a, e.lineStart = s, e.lineIndent = c;
          break;
        }
    }
    o && (wt(e, r, i, !1), Qo(e, e.line - a), r = i = e.position, o = !1), ct(f) || (i = e.position + 1), f = e.input.charCodeAt(++e.position);
  }
  return wt(e, r, i, !1), e.result ? !0 : (e.kind = p, e.result = l, !1);
}
function l0(e, t) {
  let n, r, i = e.input.charCodeAt(e.position);
  if (i !== 39)
    return !1;
  for (e.kind = "scalar", e.result = "", e.position++, n = r = e.position; (i = e.input.charCodeAt(e.position)) !== 0; )
    if (i === 39)
      if (wt(e, n, e.position, !0), i = e.input.charCodeAt(++e.position), i === 39)
        n = e.position, e.position++, r = e.position;
      else
        return !0;
    else rt(i) ? (wt(e, n, r, !0), Qo(e, le(e, !1, t)), n = r = e.position) : e.position === e.lineStart && li(e) ? x(e, "unexpected end of the document within a single quoted scalar") : (e.position++, ct(i) || (r = e.position));
  x(e, "unexpected end of the stream within a single quoted scalar");
}
function c0(e, t) {
  let n, r, i, o = e.input.charCodeAt(e.position);
  if (o !== 34)
    return !1;
  for (e.kind = "scalar", e.result = "", e.position++, n = r = e.position; (o = e.input.charCodeAt(e.position)) !== 0; ) {
    if (o === 34)
      return wt(e, n, e.position, !0), e.position++, !0;
    if (o === 92) {
      if (wt(e, n, e.position, !0), o = e.input.charCodeAt(++e.position), rt(o))
        le(e, !1, t);
      else if (o < 256 && vu[o])
        e.result += Tu[o], e.position++;
      else if ((i = e0(o)) > 0) {
        let a = i, s = 0;
        for (; a > 0; a--)
          o = e.input.charCodeAt(++e.position), (i = Zg(o)) >= 0 ? s = (s << 4) + i : x(e, "expected hexadecimal character");
        e.result += n0(s), e.position++;
      } else
        x(e, "unknown escape sequence");
      n = r = e.position;
    } else rt(o) ? (wt(e, n, r, !0), Qo(e, le(e, !1, t)), n = r = e.position) : e.position === e.lineStart && li(e) ? x(e, "unexpected end of the document within a double quoted scalar") : (e.position++, ct(o) || (r = e.position));
  }
  x(e, "unexpected end of the stream within a double quoted scalar");
}
function u0(e, t) {
  let n = !0, r, i, o;
  const a = e.tag;
  let s;
  const c = e.anchor;
  let p, l, f, h;
  const g = /* @__PURE__ */ Object.create(null);
  let _, E, T, A = e.input.charCodeAt(e.position);
  if (A === 91)
    p = 93, h = !1, s = [];
  else if (A === 123)
    p = 125, h = !0, s = {};
  else
    return !1;
  for (e.anchor !== null && kt(e, e.anchor, s), A = e.input.charCodeAt(++e.position); A !== 0; ) {
    if (le(e, !0, t), A = e.input.charCodeAt(e.position), A === p)
      return e.position++, e.tag = a, e.anchor = c, e.kind = h ? "mapping" : "sequence", e.result = s, !0;
    if (n ? A === 44 && x(e, "expected the node content, but found ','") : x(e, "missed comma between flow collection entries"), E = _ = T = null, l = f = !1, A === 63) {
      const R = e.input.charCodeAt(e.position + 1);
      Ue(R) && (l = f = !0, e.position++, le(e, !0, t));
    }
    r = e.line, i = e.lineStart, o = e.position, pn(e, t, Vr, !1, !0), E = e.tag, _ = e.result, le(e, !0, t), A = e.input.charCodeAt(e.position), (f || e.line === r) && A === 58 && (l = !0, A = e.input.charCodeAt(++e.position), le(e, !0, t), pn(e, t, Vr, !1, !0), T = e.result), h ? an(e, s, g, E, _, T, r, i, o) : l ? s.push(an(e, null, g, E, _, T, r, i, o)) : s.push(_), le(e, !0, t), A = e.input.charCodeAt(e.position), A === 44 ? (n = !0, A = e.input.charCodeAt(++e.position)) : n = !1;
  }
  x(e, "unexpected end of the stream within a flow collection");
}
function f0(e, t) {
  let n, r = Wi, i = !1, o = !1, a = t, s = 0, c = !1, p, l = e.input.charCodeAt(e.position);
  if (l === 124)
    n = !1;
  else if (l === 62)
    n = !0;
  else
    return !1;
  for (e.kind = "scalar", e.result = ""; l !== 0; )
    if (l = e.input.charCodeAt(++e.position), l === 43 || l === 45)
      Wi === r ? r = l === 43 ? gs : Xg : x(e, "repeat of a chomping mode identifier");
    else if ((p = t0(l)) >= 0)
      p === 0 ? x(e, "bad explicit indentation width of a block scalar; it cannot be less than one") : o ? x(e, "repeat of an indentation width identifier") : (a = t + p - 1, o = !0);
    else
      break;
  if (ct(l)) {
    do
      l = e.input.charCodeAt(++e.position);
    while (ct(l));
    if (l === 35)
      do
        l = e.input.charCodeAt(++e.position);
      while (!rt(l) && l !== 0);
  }
  for (; l !== 0; ) {
    for (Jo(e), e.lineIndent = 0, l = e.input.charCodeAt(e.position); (!o || e.lineIndent < a) && l === 32; )
      e.lineIndent++, l = e.input.charCodeAt(++e.position);
    if (!o && e.lineIndent > a && (a = e.lineIndent), rt(l)) {
      s++;
      continue;
    }
    if (!o && a === 0 && x(e, "missing indentation for block scalar"), e.lineIndent < a) {
      r === gs ? e.result += Ut.repeat(`
`, i ? 1 + s : s) : r === Wi && i && (e.result += `
`);
      break;
    }
    n ? ct(l) ? (c = !0, e.result += Ut.repeat(`
`, i ? 1 + s : s)) : c ? (c = !1, e.result += Ut.repeat(`
`, s + 1)) : s === 0 ? i && (e.result += " ") : e.result += Ut.repeat(`
`, s) : e.result += Ut.repeat(`
`, i ? 1 + s : s), i = !0, o = !0, s = 0;
    const f = e.position;
    for (; !rt(l) && l !== 0; )
      l = e.input.charCodeAt(++e.position);
    wt(e, f, e.position, !1);
  }
  return !0;
}
function Ts(e, t) {
  const n = e.tag, r = e.anchor, i = [];
  let o = !1;
  if (e.firstTabInLine !== -1) return !1;
  e.anchor !== null && kt(e, e.anchor, i);
  let a = e.input.charCodeAt(e.position);
  for (; a !== 0 && (e.firstTabInLine !== -1 && (e.position = e.firstTabInLine, x(e, "tab characters must not be used in indentation")), a === 45); ) {
    const s = e.input.charCodeAt(e.position + 1);
    if (!Ue(s))
      break;
    if (o = !0, e.position++, le(e, !0, -1) && e.lineIndent <= t) {
      i.push(null), a = e.input.charCodeAt(e.position);
      continue;
    }
    const c = e.line;
    if (pn(e, t, Eu, !1, !0), i.push(e.result), le(e, !0, -1), a = e.input.charCodeAt(e.position), (e.line === c || e.lineIndent > t) && a !== 0)
      x(e, "bad indentation of a sequence entry");
    else if (e.lineIndent < t)
      break;
  }
  return o ? (e.tag = n, e.anchor = r, e.kind = "sequence", e.result = i, !0) : !1;
}
function Cu(e, t, n) {
  let r, i, o, a;
  const s = e.tag, c = e.anchor, p = {}, l = /* @__PURE__ */ Object.create(null);
  let f = null, h = null, g = null, _ = !1, E = !1;
  if (e.firstTabInLine !== -1) return !1;
  e.anchor !== null && kt(e, e.anchor, p);
  let T = e.input.charCodeAt(e.position);
  for (; T !== 0; ) {
    !_ && e.firstTabInLine !== -1 && (e.position = e.firstTabInLine, x(e, "tab characters must not be used in indentation"));
    const A = e.input.charCodeAt(e.position + 1), R = e.line;
    if ((T === 63 || T === 58) && Ue(A))
      T === 63 ? (_ && (an(e, p, l, f, h, null, i, o, a), f = h = g = null), E = !0, _ = !0, r = !0) : _ ? (_ = !1, r = !0) : x(e, "incomplete explicit mapping pair; a key node is missed; or followed by a non-tabulated empty line"), e.position += 1, T = A;
    else {
      if (i = e.line, o = e.lineStart, a = e.position, !pn(e, n, gu, !1, !0))
        break;
      if (e.line === R) {
        for (T = e.input.charCodeAt(e.position); ct(T); )
          T = e.input.charCodeAt(++e.position);
        if (T === 58)
          T = e.input.charCodeAt(++e.position), Ue(T) || x(e, "a whitespace character is expected after the key-value separator within a block mapping"), _ && (an(e, p, l, f, h, null, i, o, a), f = h = g = null), E = !0, _ = !1, r = !1, f = e.tag, h = e.result;
        else if (E)
          x(e, "can not read an implicit mapping pair; a colon is missed");
        else
          return e.tag = s, e.anchor = c, !0;
      } else if (E)
        x(e, "can not read a block mapping entry; a multiline key may not be an implicit key");
      else
        return e.tag = s, e.anchor = c, !0;
    }
    if ((e.line === R || e.lineIndent > t) && (_ && (i = e.line, o = e.lineStart, a = e.position), pn(e, t, zr, !0, r) && (_ ? h = e.result : g = e.result), _ || (an(e, p, l, f, h, g, i, o, a), f = h = g = null), le(e, !0, -1), T = e.input.charCodeAt(e.position)), (e.line === R || e.lineIndent > t) && T !== 0)
      x(e, "bad indentation of a mapping entry");
    else if (e.lineIndent < t)
      break;
  }
  return _ && an(e, p, l, f, h, null, i, o, a), E && (e.tag = s, e.anchor = c, e.kind = "mapping", e.result = p), E;
}
function d0(e) {
  let t = !1, n = !1, r, i, o = e.input.charCodeAt(e.position);
  if (o !== 33) return !1;
  e.tag !== null && x(e, "duplication of a tag property"), o = e.input.charCodeAt(++e.position), o === 60 ? (t = !0, o = e.input.charCodeAt(++e.position)) : o === 33 ? (n = !0, r = "!!", o = e.input.charCodeAt(++e.position)) : r = "!";
  let a = e.position;
  if (t) {
    do
      o = e.input.charCodeAt(++e.position);
    while (o !== 0 && o !== 62);
    e.position < e.length ? (i = e.input.slice(a, e.position), o = e.input.charCodeAt(++e.position)) : x(e, "unexpected end of the stream within a verbatim tag");
  } else {
    for (; o !== 0 && !Ue(o); )
      o === 33 && (n ? x(e, "tag suffix cannot contain exclamation marks") : (r = e.input.slice(a - 1, e.position + 1), yu.test(r) || x(e, "named tag handle cannot contain such characters"), n = !0, a = e.position + 1)), o = e.input.charCodeAt(++e.position);
    i = e.input.slice(a, e.position), Qg.test(i) && x(e, "tag suffix cannot contain flow indicator characters");
  }
  i && !_u.test(i) && x(e, "tag name cannot contain such characters: " + i);
  try {
    i = decodeURIComponent(i);
  } catch {
    x(e, "tag name is malformed: " + i);
  }
  return t ? e.tag = i : Xe.call(e.tagMap, r) ? e.tag = e.tagMap[r] + i : r === "!" ? e.tag = "!" + i : r === "!!" ? e.tag = "tag:yaml.org,2002:" + i : x(e, 'undeclared tag handle "' + r + '"'), !0;
}
function h0(e) {
  let t = e.input.charCodeAt(e.position);
  if (t !== 38) return !1;
  e.anchor !== null && x(e, "duplication of an anchor property"), t = e.input.charCodeAt(++e.position);
  const n = e.position;
  for (; t !== 0 && !Ue(t) && !on(t); )
    t = e.input.charCodeAt(++e.position);
  return e.position === n && x(e, "name of an anchor node must contain at least one character"), e.anchor = e.input.slice(n, e.position), !0;
}
function p0(e) {
  let t = e.input.charCodeAt(e.position);
  if (t !== 42) return !1;
  t = e.input.charCodeAt(++e.position);
  const n = e.position;
  for (; t !== 0 && !Ue(t) && !on(t); )
    t = e.input.charCodeAt(++e.position);
  e.position === n && x(e, "name of an alias node must contain at least one character");
  const r = e.input.slice(n, e.position);
  return Xe.call(e.anchorMap, r) || x(e, 'unidentified alias "' + r + '"'), e.result = e.anchorMap[r], le(e, !0, -1), !0;
}
function m0(e, t, n, r) {
  const i = Su(e);
  return i0(e), _s(e, t), e.tag = null, e.anchor = null, e.kind = null, e.result = null, Cu(e, n, r) && e.kind === "mapping" ? (o0(e), !0) : (a0(e), _s(e, i), !1);
}
function pn(e, t, n, r, i) {
  let o, a, s = 1, c = !1, p = !1, l = null, f, h, g;
  e.depth >= e.maxDepth && x(e, "nesting exceeded maxDepth (" + e.maxDepth + ")"), e.depth += 1, e.listener !== null && e.listener("open", e), e.tag = null, e.anchor = null, e.kind = null, e.result = null;
  const _ = o = a = zr === n || Eu === n;
  if (r && le(e, !0, -1) && (c = !0, e.lineIndent > t ? s = 1 : e.lineIndent === t ? s = 0 : e.lineIndent < t && (s = -1)), s === 1)
    for (; ; ) {
      const E = e.input.charCodeAt(e.position), T = Su(e);
      if (c && (E === 33 && e.tag !== null || E === 38 && e.anchor !== null) || !d0(e) && !h0(e))
        break;
      l === null && (l = T), le(e, !0, -1) ? (c = !0, a = _, e.lineIndent > t ? s = 1 : e.lineIndent === t ? s = 0 : e.lineIndent < t && (s = -1)) : a = !1;
    }
  if (a && (a = c || i), s === 1 || zr === n)
    if (Vr === n || gu === n ? h = t : h = t + 1, g = e.position - e.lineStart, s === 1)
      if (a && (Ts(e, g) || Cu(e, g, h)) || u0(e, h))
        p = !0;
      else {
        const E = e.input.charCodeAt(e.position);
        l !== null && _ && !a && E !== 124 && E !== 62 && m0(
          e,
          l,
          l.position - l.lineStart,
          h
        ) || o && f0(e, h) || l0(e, h) || c0(e, h) ? p = !0 : p0(e) ? (p = !0, (e.tag !== null || e.anchor !== null) && x(e, "alias node should not have any properties")) : s0(e, h, Vr === n) && (p = !0, e.tag === null && (e.tag = "?")), e.anchor !== null && kt(e, e.anchor, e.result);
      }
    else s === 0 && (p = a && Ts(e, g));
  if (e.tag === null)
    e.anchor !== null && kt(e, e.anchor, e.result);
  else if (e.tag === "?") {
    e.result !== null && e.kind !== "scalar" && x(e, 'unacceptable node kind for !<?> tag; it should be "scalar", not "' + e.kind + '"');
    for (let E = 0, T = e.implicitTypes.length; E < T; E += 1)
      if (f = e.implicitTypes[E], f.resolve(e.result)) {
        e.result = f.construct(e.result), e.tag = f.tag, e.anchor !== null && kt(e, e.anchor, e.result);
        break;
      }
  } else if (e.tag !== "!") {
    if (Xe.call(e.typeMap[e.kind || "fallback"], e.tag))
      f = e.typeMap[e.kind || "fallback"][e.tag];
    else {
      f = null;
      const E = e.typeMap.multi[e.kind || "fallback"];
      for (let T = 0, A = E.length; T < A; T += 1)
        if (e.tag.slice(0, E[T].tag.length) === E[T].tag) {
          f = E[T];
          break;
        }
    }
    f || x(e, "unknown tag !<" + e.tag + ">"), e.result !== null && f.kind !== e.kind && x(e, "unacceptable node kind for !<" + e.tag + '> tag; it should be "' + f.kind + '", not "' + e.kind + '"'), f.resolve(e.result, e.tag) ? (e.result = f.construct(e.result, e.tag), e.anchor !== null && kt(e, e.anchor, e.result)) : x(e, "cannot resolve a node with !<" + e.tag + "> explicit tag");
  }
  return e.listener !== null && e.listener("close", e), e.depth -= 1, e.tag !== null || e.anchor !== null || p;
}
function g0(e) {
  const t = e.position;
  let n = !1, r;
  for (e.version = null, e.checkLineBreaks = e.legacy, e.tagMap = /* @__PURE__ */ Object.create(null), e.anchorMap = /* @__PURE__ */ Object.create(null); (r = e.input.charCodeAt(e.position)) !== 0 && (le(e, !0, -1), r = e.input.charCodeAt(e.position), !(e.lineIndent > 0 || r !== 37)); ) {
    n = !0, r = e.input.charCodeAt(++e.position);
    let i = e.position;
    for (; r !== 0 && !Ue(r); )
      r = e.input.charCodeAt(++e.position);
    const o = e.input.slice(i, e.position), a = [];
    for (o.length < 1 && x(e, "directive name must not be less than one character in length"); r !== 0; ) {
      for (; ct(r); )
        r = e.input.charCodeAt(++e.position);
      if (r === 35) {
        do
          r = e.input.charCodeAt(++e.position);
        while (r !== 0 && !rt(r));
        break;
      }
      if (rt(r)) break;
      for (i = e.position; r !== 0 && !Ue(r); )
        r = e.input.charCodeAt(++e.position);
      a.push(e.input.slice(i, e.position));
    }
    r !== 0 && Jo(e), Xe.call(ws, o) ? ws[o](e, o, a) : Yr(e, 'unknown document directive "' + o + '"');
  }
  if (le(e, !0, -1), e.lineIndent === 0 && e.input.charCodeAt(e.position) === 45 && e.input.charCodeAt(e.position + 1) === 45 && e.input.charCodeAt(e.position + 2) === 45 ? (e.position += 3, le(e, !0, -1)) : n && x(e, "directives end mark is expected"), pn(e, e.lineIndent - 1, zr, !1, !0), le(e, !0, -1), e.checkLineBreaks && Jg.test(e.input.slice(t, e.position)) && Yr(e, "non-ASCII line breaks are interpreted as content"), e.documents.push(e.result), e.position === e.lineStart && li(e)) {
    e.input.charCodeAt(e.position) === 46 && (e.position += 3, le(e, !0, -1));
    return;
  }
  e.position < e.length - 1 && x(e, "end of the stream or a document separator is expected");
}
function bu(e, t) {
  e = String(e), t = t || {}, e.length !== 0 && (e.charCodeAt(e.length - 1) !== 10 && e.charCodeAt(e.length - 1) !== 13 && (e += `
`), e.charCodeAt(0) === 65279 && (e = e.slice(1)));
  const n = new r0(e, t), r = e.indexOf("\0");
  for (r !== -1 && (n.position = r, x(n, "null byte is not allowed in input")), n.input += "\0"; n.input.charCodeAt(n.position) === 32; )
    n.lineIndent += 1, n.position += 1;
  for (; n.position < n.length - 1; )
    g0(n);
  return n.documents;
}
function E0(e, t, n) {
  t !== null && typeof t == "object" && typeof n > "u" && (n = t, t = null);
  const r = bu(e, n);
  if (typeof t != "function")
    return r;
  for (let i = 0, o = r.length; i < o; i += 1)
    t(r[i]);
}
function y0(e, t) {
  const n = bu(e, t);
  if (n.length !== 0) {
    if (n.length === 1)
      return n[0];
    throw new mu("expected a single document in the stream, but found more");
  }
}
Yo.loadAll = E0;
Yo.load = y0;
var Nu = {};
const ci = Ke, ar = or, _0 = Ko, Ru = Object.prototype.toString, $u = Object.prototype.hasOwnProperty, Zo = 65279, w0 = 9, qn = 10, v0 = 13, T0 = 32, A0 = 33, S0 = 34, Co = 35, C0 = 37, b0 = 38, N0 = 39, R0 = 42, Ou = 44, $0 = 45, Xr = 58, O0 = 61, I0 = 62, D0 = 63, P0 = 64, Iu = 91, Du = 93, L0 = 96, Pu = 123, F0 = 124, Lu = 125, Ne = {};
Ne[0] = "\\0";
Ne[7] = "\\a";
Ne[8] = "\\b";
Ne[9] = "\\t";
Ne[10] = "\\n";
Ne[11] = "\\v";
Ne[12] = "\\f";
Ne[13] = "\\r";
Ne[27] = "\\e";
Ne[34] = '\\"';
Ne[92] = "\\\\";
Ne[133] = "\\N";
Ne[160] = "\\_";
Ne[8232] = "\\L";
Ne[8233] = "\\P";
const x0 = [
  "y",
  "Y",
  "yes",
  "Yes",
  "YES",
  "on",
  "On",
  "ON",
  "n",
  "N",
  "no",
  "No",
  "NO",
  "off",
  "Off",
  "OFF"
], U0 = /^[-+]?[0-9_]+(?::[0-9_]+)+(?:\.[0-9_]*)?$/;
function k0(e, t) {
  if (t === null) return {};
  const n = {}, r = Object.keys(t);
  for (let i = 0, o = r.length; i < o; i += 1) {
    let a = r[i], s = String(t[a]);
    a.slice(0, 2) === "!!" && (a = "tag:yaml.org,2002:" + a.slice(2));
    const c = e.compiledTypeMap.fallback[a];
    c && $u.call(c.styleAliases, s) && (s = c.styleAliases[s]), n[a] = s;
  }
  return n;
}
function M0(e) {
  let t, n;
  const r = e.toString(16).toUpperCase();
  if (e <= 255)
    t = "x", n = 2;
  else if (e <= 65535)
    t = "u", n = 4;
  else if (e <= 4294967295)
    t = "U", n = 8;
  else
    throw new ar("code point within a string may not be greater than 0xFFFFFFFF");
  return "\\" + t + ci.repeat("0", n - r.length) + r;
}
const B0 = 1, Wn = 2;
function H0(e) {
  this.schema = e.schema || _0, this.indent = Math.max(1, e.indent || 2), this.noArrayIndent = e.noArrayIndent || !1, this.skipInvalid = e.skipInvalid || !1, this.flowLevel = ci.isNothing(e.flowLevel) ? -1 : e.flowLevel, this.styleMap = k0(this.schema, e.styles || null), this.sortKeys = e.sortKeys || !1, this.lineWidth = e.lineWidth || 80, this.noRefs = e.noRefs || !1, this.noCompatMode = e.noCompatMode || !1, this.condenseFlow = e.condenseFlow || !1, this.quotingType = e.quotingType === '"' ? Wn : B0, this.forceQuotes = e.forceQuotes || !1, this.replacer = typeof e.replacer == "function" ? e.replacer : null, this.implicitTypes = this.schema.compiledImplicit, this.explicitTypes = this.schema.compiledExplicit, this.tag = null, this.result = "", this.duplicates = [], this.usedDuplicates = null;
}
function As(e, t) {
  const n = ci.repeat(" ", t);
  let r = 0, i = "";
  const o = e.length;
  for (; r < o; ) {
    let a;
    const s = e.indexOf(`
`, r);
    s === -1 ? (a = e.slice(r), r = o) : (a = e.slice(r, s + 1), r = s + 1), a.length && a !== `
` && (i += n), i += a;
  }
  return i;
}
function bo(e, t) {
  return `
` + ci.repeat(" ", e.indent * t);
}
function j0(e, t) {
  for (let n = 0, r = e.implicitTypes.length; n < r; n += 1)
    if (e.implicitTypes[n].resolve(t))
      return !0;
  return !1;
}
function Kr(e) {
  return e === T0 || e === w0;
}
function Vn(e) {
  return e >= 32 && e <= 126 || e >= 161 && e <= 55295 && e !== 8232 && e !== 8233 || e >= 57344 && e <= 65533 && e !== Zo || e >= 65536 && e <= 1114111;
}
function Ss(e) {
  return Vn(e) && e !== Zo && // - b-char
  e !== v0 && e !== qn;
}
function Cs(e, t, n) {
  const r = Ss(e), i = r && !Kr(e);
  return (
    // ns-plain-safe
    (n ? r : r && // - c-flow-indicator
    e !== Ou && e !== Iu && e !== Du && e !== Pu && e !== Lu) && // ns-plain-char
    e !== Co && // false on '#'
    !(t === Xr && !i) || // false on ': '
    Ss(t) && !Kr(t) && e === Co || // change to true on '[^ ]#'
    t === Xr && i
  );
}
function G0(e) {
  return Vn(e) && e !== Zo && !Kr(e) && // - s-white
  // - (c-indicator ::=
  // “-” | “?” | “:” | “,” | “[” | “]” | “{” | “}”
  e !== $0 && e !== D0 && e !== Xr && e !== Ou && e !== Iu && e !== Du && e !== Pu && e !== Lu && // | “#” | “&” | “*” | “!” | “|” | “=” | “>” | “'” | “"”
  e !== Co && e !== b0 && e !== R0 && e !== A0 && e !== F0 && e !== O0 && e !== I0 && e !== N0 && e !== S0 && // | “%” | “@” | “`”)
  e !== C0 && e !== P0 && e !== L0;
}
function q0(e) {
  return !Kr(e) && e !== Xr;
}
function Dn(e, t) {
  const n = e.charCodeAt(t);
  let r;
  return n >= 55296 && n <= 56319 && t + 1 < e.length && (r = e.charCodeAt(t + 1), r >= 56320 && r <= 57343) ? (n - 55296) * 1024 + r - 56320 + 65536 : n;
}
function Fu(e) {
  return /^\n* /.test(e);
}
const xu = 1, No = 2, Uu = 3, ku = 4, tn = 5;
function W0(e, t, n, r, i, o, a, s) {
  let c, p = 0, l = null, f = !1, h = !1;
  const g = r !== -1;
  let _ = -1, E = G0(Dn(e, 0)) && q0(Dn(e, e.length - 1));
  if (t || a)
    for (c = 0; c < e.length; p >= 65536 ? c += 2 : c++) {
      if (p = Dn(e, c), !Vn(p))
        return tn;
      E = E && Cs(p, l, s), l = p;
    }
  else {
    for (c = 0; c < e.length; p >= 65536 ? c += 2 : c++) {
      if (p = Dn(e, c), p === qn)
        f = !0, g && (h = h || // Foldable line = too long, and not more-indented.
        c - _ - 1 > r && e[_ + 1] !== " ", _ = c);
      else if (!Vn(p))
        return tn;
      E = E && Cs(p, l, s), l = p;
    }
    h = h || g && c - _ - 1 > r && e[_ + 1] !== " ";
  }
  return !f && !h ? E && !a && !i(e) ? xu : o === Wn ? tn : No : n > 9 && Fu(e) ? tn : a ? o === Wn ? tn : No : h ? ku : Uu;
}
function V0(e, t, n, r, i) {
  e.dump = function() {
    if (t.length === 0)
      return e.quotingType === Wn ? '""' : "''";
    if (!e.noCompatMode && (x0.indexOf(t) !== -1 || U0.test(t)))
      return e.quotingType === Wn ? '"' + t + '"' : "'" + t + "'";
    const o = e.indent * Math.max(1, n), a = e.lineWidth === -1 ? -1 : Math.max(Math.min(e.lineWidth, 40), e.lineWidth - o), s = r || // No block styles in flow mode.
    e.flowLevel > -1 && n >= e.flowLevel;
    function c(p) {
      return j0(e, p);
    }
    switch (W0(
      t,
      s,
      e.indent,
      a,
      c,
      e.quotingType,
      e.forceQuotes && !r,
      i
    )) {
      case xu:
        return t;
      case No:
        return "'" + t.replace(/'/g, "''") + "'";
      case Uu:
        return "|" + bs(t, e.indent) + Ns(As(t, o));
      case ku:
        return ">" + bs(t, e.indent) + Ns(As(z0(t, a), o));
      case tn:
        return '"' + Y0(t) + '"';
      default:
        throw new ar("impossible error: invalid scalar style");
    }
  }();
}
function bs(e, t) {
  const n = Fu(e) ? String(t) : "", r = e[e.length - 1] === `
`, o = r && (e[e.length - 2] === `
` || e === `
`) ? "+" : r ? "" : "-";
  return n + o + `
`;
}
function Ns(e) {
  return e[e.length - 1] === `
` ? e.slice(0, -1) : e;
}
function z0(e, t) {
  const n = /(\n+)([^\n]*)/g;
  let r = function() {
    let s = e.indexOf(`
`);
    return s = s !== -1 ? s : e.length, n.lastIndex = s, Rs(e.slice(0, s), t);
  }(), i = e[0] === `
` || e[0] === " ", o, a;
  for (; a = n.exec(e); ) {
    const s = a[1], c = a[2];
    o = c[0] === " ", r += s + (!i && !o && c !== "" ? `
` : "") + Rs(c, t), i = o;
  }
  return r;
}
function Rs(e, t) {
  if (e === "" || e[0] === " ") return e;
  const n = / [^ ]/g;
  let r, i = 0, o, a = 0, s = 0, c = "";
  for (; r = n.exec(e); )
    s = r.index, s - i > t && (o = a > i ? a : s, c += `
` + e.slice(i, o), i = o + 1), a = s;
  return c += `
`, e.length - i > t && a > i ? c += e.slice(i, a) + `
` + e.slice(a + 1) : c += e.slice(i), c.slice(1);
}
function Y0(e) {
  let t = "", n = 0;
  for (let r = 0; r < e.length; n >= 65536 ? r += 2 : r++) {
    n = Dn(e, r);
    const i = Ne[n];
    !i && Vn(n) ? (t += e[r], n >= 65536 && (t += e[r + 1])) : t += i || M0(n);
  }
  return t;
}
function X0(e, t, n) {
  let r = "";
  const i = e.tag;
  for (let o = 0, a = n.length; o < a; o += 1) {
    let s = n[o];
    e.replacer && (s = e.replacer.call(n, String(o), s)), (ut(e, t, s, !1, !1) || typeof s > "u" && ut(e, t, null, !1, !1)) && (r !== "" && (r += "," + (e.condenseFlow ? "" : " ")), r += e.dump);
  }
  e.tag = i, e.dump = "[" + r + "]";
}
function $s(e, t, n, r) {
  let i = "";
  const o = e.tag;
  for (let a = 0, s = n.length; a < s; a += 1) {
    let c = n[a];
    e.replacer && (c = e.replacer.call(n, String(a), c)), (ut(e, t + 1, c, !0, !0, !1, !0) || typeof c > "u" && ut(e, t + 1, null, !0, !0, !1, !0)) && ((!r || i !== "") && (i += bo(e, t)), e.dump && qn === e.dump.charCodeAt(0) ? i += "-" : i += "- ", i += e.dump);
  }
  e.tag = o, e.dump = i || "[]";
}
function K0(e, t, n) {
  let r = "";
  const i = e.tag, o = Object.keys(n);
  for (let a = 0, s = o.length; a < s; a += 1) {
    let c = "";
    r !== "" && (c += ", "), e.condenseFlow && (c += '"');
    const p = o[a];
    let l = n[p];
    e.replacer && (l = e.replacer.call(n, p, l)), ut(e, t, p, !1, !1) && (e.dump.length > 1024 && (c += "? "), c += e.dump + (e.condenseFlow ? '"' : "") + ":" + (e.condenseFlow ? "" : " "), ut(e, t, l, !1, !1) && (c += e.dump, r += c));
  }
  e.tag = i, e.dump = "{" + r + "}";
}
function J0(e, t, n, r) {
  let i = "";
  const o = e.tag, a = Object.keys(n);
  if (e.sortKeys === !0)
    a.sort();
  else if (typeof e.sortKeys == "function")
    a.sort(e.sortKeys);
  else if (e.sortKeys)
    throw new ar("sortKeys must be a boolean or a function");
  for (let s = 0, c = a.length; s < c; s += 1) {
    let p = "";
    (!r || i !== "") && (p += bo(e, t));
    const l = a[s];
    let f = n[l];
    if (e.replacer && (f = e.replacer.call(n, l, f)), !ut(e, t + 1, l, !0, !0, !0))
      continue;
    const h = e.tag !== null && e.tag !== "?" || e.dump && e.dump.length > 1024;
    h && (e.dump && qn === e.dump.charCodeAt(0) ? p += "?" : p += "? "), p += e.dump, h && (p += bo(e, t)), ut(e, t + 1, f, !0, h) && (e.dump && qn === e.dump.charCodeAt(0) ? p += ":" : p += ": ", p += e.dump, i += p);
  }
  e.tag = o, e.dump = i || "{}";
}
function Os(e, t, n) {
  const r = n ? e.explicitTypes : e.implicitTypes;
  for (let i = 0, o = r.length; i < o; i += 1) {
    const a = r[i];
    if ((a.instanceOf || a.predicate) && (!a.instanceOf || typeof t == "object" && t instanceof a.instanceOf) && (!a.predicate || a.predicate(t))) {
      if (n ? a.multi && a.representName ? e.tag = a.representName(t) : e.tag = a.tag : e.tag = "?", a.represent) {
        const s = e.styleMap[a.tag] || a.defaultStyle;
        let c;
        if (Ru.call(a.represent) === "[object Function]")
          c = a.represent(t, s);
        else if ($u.call(a.represent, s))
          c = a.represent[s](t, s);
        else
          throw new ar("!<" + a.tag + '> tag resolver accepts not "' + s + '" style');
        e.dump = c;
      }
      return !0;
    }
  }
  return !1;
}
function ut(e, t, n, r, i, o, a) {
  e.tag = null, e.dump = n, Os(e, n, !1) || Os(e, n, !0);
  const s = Ru.call(e.dump), c = r;
  r && (r = e.flowLevel < 0 || e.flowLevel > t);
  const p = s === "[object Object]" || s === "[object Array]";
  let l, f;
  if (p && (l = e.duplicates.indexOf(n), f = l !== -1), (e.tag !== null && e.tag !== "?" || f || e.indent !== 2 && t > 0) && (i = !1), f && e.usedDuplicates[l])
    e.dump = "*ref_" + l;
  else {
    if (p && f && !e.usedDuplicates[l] && (e.usedDuplicates[l] = !0), s === "[object Object]")
      r && Object.keys(e.dump).length !== 0 ? (J0(e, t, e.dump, i), f && (e.dump = "&ref_" + l + e.dump)) : (K0(e, t, e.dump), f && (e.dump = "&ref_" + l + " " + e.dump));
    else if (s === "[object Array]")
      r && e.dump.length !== 0 ? (e.noArrayIndent && !a && t > 0 ? $s(e, t - 1, e.dump, i) : $s(e, t, e.dump, i), f && (e.dump = "&ref_" + l + e.dump)) : (X0(e, t, e.dump), f && (e.dump = "&ref_" + l + " " + e.dump));
    else if (s === "[object String]")
      e.tag !== "?" && V0(e, e.dump, t, o, c);
    else {
      if (s === "[object Undefined]")
        return !1;
      if (e.skipInvalid) return !1;
      throw new ar("unacceptable kind of an object to dump " + s);
    }
    if (e.tag !== null && e.tag !== "?") {
      let h = encodeURI(
        e.tag[0] === "!" ? e.tag.slice(1) : e.tag
      ).replace(/!/g, "%21");
      e.tag[0] === "!" ? h = "!" + h : h.slice(0, 18) === "tag:yaml.org,2002:" ? h = "!!" + h.slice(18) : h = "!<" + h + ">", e.dump = h + " " + e.dump;
    }
  }
  return !0;
}
function Q0(e, t) {
  const n = [], r = [];
  Ro(e, n, r);
  const i = r.length;
  for (let o = 0; o < i; o += 1)
    t.duplicates.push(n[r[o]]);
  t.usedDuplicates = new Array(i);
}
function Ro(e, t, n) {
  if (e !== null && typeof e == "object") {
    const r = t.indexOf(e);
    if (r !== -1)
      n.indexOf(r) === -1 && n.push(r);
    else if (t.push(e), Array.isArray(e))
      for (let i = 0, o = e.length; i < o; i += 1)
        Ro(e[i], t, n);
    else {
      const i = Object.keys(e);
      for (let o = 0, a = i.length; o < a; o += 1)
        Ro(e[i[o]], t, n);
    }
  }
}
function Z0(e, t) {
  t = t || {};
  const n = new H0(t);
  n.noRefs || Q0(e, n);
  let r = e;
  return n.replacer && (r = n.replacer.call({ "": r }, "", r)), ut(n, 0, r, !0, !0) ? n.dump + `
` : "";
}
Nu.dump = Z0;
const Mu = Yo, eE = Nu;
function ea(e, t) {
  return function() {
    throw new Error("Function yaml." + e + " is removed in js-yaml 4. Use yaml." + t + " instead, which is now safe by default.");
  };
}
we.Type = Le;
we.Schema = Xc;
we.FAILSAFE_SCHEMA = Zc;
we.JSON_SCHEMA = ou;
we.CORE_SCHEMA = au;
we.DEFAULT_SCHEMA = Ko;
we.load = Mu.load;
we.loadAll = Mu.loadAll;
we.dump = eE.dump;
we.YAMLException = or;
we.types = {
  binary: fu,
  float: iu,
  map: Qc,
  null: eu,
  pairs: hu,
  set: pu,
  timestamp: cu,
  bool: tu,
  int: nu,
  merge: uu,
  omap: du,
  seq: Jc,
  str: Kc
};
we.safeLoad = ea("safeLoad", "load");
we.safeLoadAll = ea("safeLoadAll", "loadAll");
we.safeDump = ea("safeDump", "dump");
var ui = {};
Object.defineProperty(ui, "__esModule", { value: !0 });
ui.Lazy = void 0;
class tE {
  constructor(t) {
    this._value = null, this.creator = t;
  }
  get hasValue() {
    return this.creator == null;
  }
  get value() {
    if (this.creator == null)
      return this._value;
    const t = this.creator();
    return this.value = t, t;
  }
  set value(t) {
    this._value = t, this.creator = null;
  }
}
ui.Lazy = tE;
var $o = { exports: {} };
const nE = "2.0.0", Bu = 256, rE = Number.MAX_SAFE_INTEGER || /* istanbul ignore next */
9007199254740991, iE = 16, oE = Bu - 6, aE = [
  "major",
  "premajor",
  "minor",
  "preminor",
  "patch",
  "prepatch",
  "prerelease"
];
var fi = {
  MAX_LENGTH: Bu,
  MAX_SAFE_COMPONENT_LENGTH: iE,
  MAX_SAFE_BUILD_LENGTH: oE,
  MAX_SAFE_INTEGER: rE,
  RELEASE_TYPES: aE,
  SEMVER_SPEC_VERSION: nE,
  FLAG_INCLUDE_PRERELEASE: 1,
  FLAG_LOOSE: 2
};
const sE = typeof process == "object" && process.env && process.env.NODE_DEBUG && /\bsemver\b/i.test(process.env.NODE_DEBUG) ? (...e) => console.error("SEMVER", ...e) : () => {
};
var di = sE;
(function(e, t) {
  const {
    MAX_SAFE_COMPONENT_LENGTH: n,
    MAX_SAFE_BUILD_LENGTH: r,
    MAX_LENGTH: i
  } = fi, o = di;
  t = e.exports = {};
  const a = t.re = [], s = t.safeRe = [], c = t.src = [], p = t.safeSrc = [], l = t.t = {};
  let f = 0;
  const h = "[a-zA-Z0-9-]", g = [
    ["\\s", 1],
    ["\\d", i],
    [h, r]
  ], _ = (T) => {
    for (const [A, R] of g)
      T = T.split(`${A}*`).join(`${A}{0,${R}}`).split(`${A}+`).join(`${A}{1,${R}}`);
    return T;
  }, E = (T, A, R) => {
    const P = _(A), B = f++;
    o(T, B, A), l[T] = B, c[B] = A, p[B] = P, a[B] = new RegExp(A, R ? "g" : void 0), s[B] = new RegExp(P, R ? "g" : void 0);
  };
  E("NUMERICIDENTIFIER", "0|[1-9]\\d*"), E("NUMERICIDENTIFIERLOOSE", "\\d+"), E("NONNUMERICIDENTIFIER", `\\d*[a-zA-Z-]${h}*`), E("MAINVERSION", `(${c[l.NUMERICIDENTIFIER]})\\.(${c[l.NUMERICIDENTIFIER]})\\.(${c[l.NUMERICIDENTIFIER]})`), E("MAINVERSIONLOOSE", `(${c[l.NUMERICIDENTIFIERLOOSE]})\\.(${c[l.NUMERICIDENTIFIERLOOSE]})\\.(${c[l.NUMERICIDENTIFIERLOOSE]})`), E("PRERELEASEIDENTIFIER", `(?:${c[l.NONNUMERICIDENTIFIER]}|${c[l.NUMERICIDENTIFIER]})`), E("PRERELEASEIDENTIFIERLOOSE", `(?:${c[l.NONNUMERICIDENTIFIER]}|${c[l.NUMERICIDENTIFIERLOOSE]})`), E("PRERELEASE", `(?:-(${c[l.PRERELEASEIDENTIFIER]}(?:\\.${c[l.PRERELEASEIDENTIFIER]})*))`), E("PRERELEASELOOSE", `(?:-?(${c[l.PRERELEASEIDENTIFIERLOOSE]}(?:\\.${c[l.PRERELEASEIDENTIFIERLOOSE]})*))`), E("BUILDIDENTIFIER", `${h}+`), E("BUILD", `(?:\\+(${c[l.BUILDIDENTIFIER]}(?:\\.${c[l.BUILDIDENTIFIER]})*))`), E("FULLPLAIN", `v?${c[l.MAINVERSION]}${c[l.PRERELEASE]}?${c[l.BUILD]}?`), E("FULL", `^${c[l.FULLPLAIN]}$`), E("LOOSEPLAIN", `[v=\\s]*${c[l.MAINVERSIONLOOSE]}${c[l.PRERELEASELOOSE]}?${c[l.BUILD]}?`), E("LOOSE", `^${c[l.LOOSEPLAIN]}$`), E("GTLT", "((?:<|>)?=?)"), E("XRANGEIDENTIFIERLOOSE", `${c[l.NUMERICIDENTIFIERLOOSE]}|x|X|\\*`), E("XRANGEIDENTIFIER", `${c[l.NUMERICIDENTIFIER]}|x|X|\\*`), E("XRANGEPLAIN", `[v=\\s]*(${c[l.XRANGEIDENTIFIER]})(?:\\.(${c[l.XRANGEIDENTIFIER]})(?:\\.(${c[l.XRANGEIDENTIFIER]})(?:${c[l.PRERELEASE]})?${c[l.BUILD]}?)?)?`), E("XRANGEPLAINLOOSE", `[v=\\s]*(${c[l.XRANGEIDENTIFIERLOOSE]})(?:\\.(${c[l.XRANGEIDENTIFIERLOOSE]})(?:\\.(${c[l.XRANGEIDENTIFIERLOOSE]})(?:${c[l.PRERELEASELOOSE]})?${c[l.BUILD]}?)?)?`), E("XRANGE", `^${c[l.GTLT]}\\s*${c[l.XRANGEPLAIN]}$`), E("XRANGELOOSE", `^${c[l.GTLT]}\\s*${c[l.XRANGEPLAINLOOSE]}$`), E("COERCEPLAIN", `(^|[^\\d])(\\d{1,${n}})(?:\\.(\\d{1,${n}}))?(?:\\.(\\d{1,${n}}))?`), E("COERCE", `${c[l.COERCEPLAIN]}(?:$|[^\\d])`), E("COERCEFULL", c[l.COERCEPLAIN] + `(?:${c[l.PRERELEASE]})?(?:${c[l.BUILD]})?(?:$|[^\\d])`), E("COERCERTL", c[l.COERCE], !0), E("COERCERTLFULL", c[l.COERCEFULL], !0), E("LONETILDE", "(?:~>?)"), E("TILDETRIM", `(\\s*)${c[l.LONETILDE]}\\s+`, !0), t.tildeTrimReplace = "$1~", E("TILDE", `^${c[l.LONETILDE]}${c[l.XRANGEPLAIN]}$`), E("TILDELOOSE", `^${c[l.LONETILDE]}${c[l.XRANGEPLAINLOOSE]}$`), E("LONECARET", "(?:\\^)"), E("CARETTRIM", `(\\s*)${c[l.LONECARET]}\\s+`, !0), t.caretTrimReplace = "$1^", E("CARET", `^${c[l.LONECARET]}${c[l.XRANGEPLAIN]}$`), E("CARETLOOSE", `^${c[l.LONECARET]}${c[l.XRANGEPLAINLOOSE]}$`), E("COMPARATORLOOSE", `^${c[l.GTLT]}\\s*(${c[l.LOOSEPLAIN]})$|^$`), E("COMPARATOR", `^${c[l.GTLT]}\\s*(${c[l.FULLPLAIN]})$|^$`), E("COMPARATORTRIM", `(\\s*)${c[l.GTLT]}\\s*(${c[l.LOOSEPLAIN]}|${c[l.XRANGEPLAIN]})`, !0), t.comparatorTrimReplace = "$1$2$3", E("HYPHENRANGE", `^\\s*(${c[l.XRANGEPLAIN]})\\s+-\\s+(${c[l.XRANGEPLAIN]})\\s*$`), E("HYPHENRANGELOOSE", `^\\s*(${c[l.XRANGEPLAINLOOSE]})\\s+-\\s+(${c[l.XRANGEPLAINLOOSE]})\\s*$`), E("STAR", "(<|>)?=?\\s*\\*"), E("GTE0", "^\\s*>=\\s*0\\.0\\.0\\s*$"), E("GTE0PRE", "^\\s*>=\\s*0\\.0\\.0-0\\s*$");
})($o, $o.exports);
var sr = $o.exports;
const lE = Object.freeze({ loose: !0 }), cE = Object.freeze({}), uE = (e) => e ? typeof e != "object" ? lE : e : cE;
var ta = uE;
const Is = /^[0-9]+$/, Hu = (e, t) => {
  if (typeof e == "number" && typeof t == "number")
    return e === t ? 0 : e < t ? -1 : 1;
  const n = Is.test(e), r = Is.test(t);
  return n && r && (e = +e, t = +t), e === t ? 0 : n && !r ? -1 : r && !n ? 1 : e < t ? -1 : 1;
}, fE = (e, t) => Hu(t, e);
var ju = {
  compareIdentifiers: Hu,
  rcompareIdentifiers: fE
};
const Nr = di, { MAX_LENGTH: Ds, MAX_SAFE_INTEGER: Rr } = fi, { safeRe: $r, t: Or } = sr, dE = ta, { compareIdentifiers: Vi } = ju;
let hE = class tt {
  constructor(t, n) {
    if (n = dE(n), t instanceof tt) {
      if (t.loose === !!n.loose && t.includePrerelease === !!n.includePrerelease)
        return t;
      t = t.version;
    } else if (typeof t != "string")
      throw new TypeError(`Invalid version. Must be a string. Got type "${typeof t}".`);
    if (t.length > Ds)
      throw new TypeError(
        `version is longer than ${Ds} characters`
      );
    Nr("SemVer", t, n), this.options = n, this.loose = !!n.loose, this.includePrerelease = !!n.includePrerelease;
    const r = t.trim().match(n.loose ? $r[Or.LOOSE] : $r[Or.FULL]);
    if (!r)
      throw new TypeError(`Invalid Version: ${t}`);
    if (this.raw = t, this.major = +r[1], this.minor = +r[2], this.patch = +r[3], this.major > Rr || this.major < 0)
      throw new TypeError("Invalid major version");
    if (this.minor > Rr || this.minor < 0)
      throw new TypeError("Invalid minor version");
    if (this.patch > Rr || this.patch < 0)
      throw new TypeError("Invalid patch version");
    r[4] ? this.prerelease = r[4].split(".").map((i) => {
      if (/^[0-9]+$/.test(i)) {
        const o = +i;
        if (o >= 0 && o < Rr)
          return o;
      }
      return i;
    }) : this.prerelease = [], this.build = r[5] ? r[5].split(".") : [], this.format();
  }
  format() {
    return this.version = `${this.major}.${this.minor}.${this.patch}`, this.prerelease.length && (this.version += `-${this.prerelease.join(".")}`), this.version;
  }
  toString() {
    return this.version;
  }
  compare(t) {
    if (Nr("SemVer.compare", this.version, this.options, t), !(t instanceof tt)) {
      if (typeof t == "string" && t === this.version)
        return 0;
      t = new tt(t, this.options);
    }
    return t.version === this.version ? 0 : this.compareMain(t) || this.comparePre(t);
  }
  compareMain(t) {
    return t instanceof tt || (t = new tt(t, this.options)), this.major < t.major ? -1 : this.major > t.major ? 1 : this.minor < t.minor ? -1 : this.minor > t.minor ? 1 : this.patch < t.patch ? -1 : this.patch > t.patch ? 1 : 0;
  }
  comparePre(t) {
    if (t instanceof tt || (t = new tt(t, this.options)), this.prerelease.length && !t.prerelease.length)
      return -1;
    if (!this.prerelease.length && t.prerelease.length)
      return 1;
    if (!this.prerelease.length && !t.prerelease.length)
      return 0;
    let n = 0;
    do {
      const r = this.prerelease[n], i = t.prerelease[n];
      if (Nr("prerelease compare", n, r, i), r === void 0 && i === void 0)
        return 0;
      if (i === void 0)
        return 1;
      if (r === void 0)
        return -1;
      if (r === i)
        continue;
      return Vi(r, i);
    } while (++n);
  }
  compareBuild(t) {
    t instanceof tt || (t = new tt(t, this.options));
    let n = 0;
    do {
      const r = this.build[n], i = t.build[n];
      if (Nr("build compare", n, r, i), r === void 0 && i === void 0)
        return 0;
      if (i === void 0)
        return 1;
      if (r === void 0)
        return -1;
      if (r === i)
        continue;
      return Vi(r, i);
    } while (++n);
  }
  // preminor will bump the version up to the next minor release, and immediately
  // down to pre-release. premajor and prepatch work the same way.
  inc(t, n, r) {
    if (t.startsWith("pre")) {
      if (!n && r === !1)
        throw new Error("invalid increment argument: identifier is empty");
      if (n) {
        const i = `-${n}`.match(this.options.loose ? $r[Or.PRERELEASELOOSE] : $r[Or.PRERELEASE]);
        if (!i || i[1] !== n)
          throw new Error(`invalid identifier: ${n}`);
      }
    }
    switch (t) {
      case "premajor":
        this.prerelease.length = 0, this.patch = 0, this.minor = 0, this.major++, this.inc("pre", n, r);
        break;
      case "preminor":
        this.prerelease.length = 0, this.patch = 0, this.minor++, this.inc("pre", n, r);
        break;
      case "prepatch":
        this.prerelease.length = 0, this.inc("patch", n, r), this.inc("pre", n, r);
        break;
      case "prerelease":
        this.prerelease.length === 0 && this.inc("patch", n, r), this.inc("pre", n, r);
        break;
      case "release":
        if (this.prerelease.length === 0)
          throw new Error(`version ${this.raw} is not a prerelease`);
        this.prerelease.length = 0;
        break;
      case "major":
        (this.minor !== 0 || this.patch !== 0 || this.prerelease.length === 0) && this.major++, this.minor = 0, this.patch = 0, this.prerelease = [];
        break;
      case "minor":
        (this.patch !== 0 || this.prerelease.length === 0) && this.minor++, this.patch = 0, this.prerelease = [];
        break;
      case "patch":
        this.prerelease.length === 0 && this.patch++, this.prerelease = [];
        break;
      case "pre": {
        const i = Number(r) ? 1 : 0;
        if (this.prerelease.length === 0)
          this.prerelease = [i];
        else {
          let o = this.prerelease.length;
          for (; --o >= 0; )
            typeof this.prerelease[o] == "number" && (this.prerelease[o]++, o = -2);
          if (o === -1) {
            if (n === this.prerelease.join(".") && r === !1)
              throw new Error("invalid increment argument: identifier already exists");
            this.prerelease.push(i);
          }
        }
        if (n) {
          let o = [n, i];
          r === !1 && (o = [n]), Vi(this.prerelease[0], n) === 0 ? isNaN(this.prerelease[1]) && (this.prerelease = o) : this.prerelease = o;
        }
        break;
      }
      default:
        throw new Error(`invalid increment argument: ${t}`);
    }
    return this.raw = this.format(), this.build.length && (this.raw += `+${this.build.join(".")}`), this;
  }
};
var Fe = hE;
const Ps = Fe, pE = (e, t, n = !1) => {
  if (e instanceof Ps)
    return e;
  try {
    return new Ps(e, t);
  } catch (r) {
    if (!n)
      return null;
    throw r;
  }
};
var En = pE;
const mE = En, gE = (e, t) => {
  const n = mE(e, t);
  return n ? n.version : null;
};
var EE = gE;
const yE = En, _E = (e, t) => {
  const n = yE(e.trim().replace(/^[=v]+/, ""), t);
  return n ? n.version : null;
};
var wE = _E;
const Ls = Fe, vE = (e, t, n, r, i) => {
  typeof n == "string" && (i = r, r = n, n = void 0);
  try {
    return new Ls(
      e instanceof Ls ? e.version : e,
      n
    ).inc(t, r, i).version;
  } catch {
    return null;
  }
};
var TE = vE;
const Fs = En, AE = (e, t) => {
  const n = Fs(e, null, !0), r = Fs(t, null, !0), i = n.compare(r);
  if (i === 0)
    return null;
  const o = i > 0, a = o ? n : r, s = o ? r : n, c = !!a.prerelease.length;
  if (!!s.prerelease.length && !c) {
    if (!s.patch && !s.minor)
      return "major";
    if (s.compareMain(a) === 0)
      return s.minor && !s.patch ? "minor" : "patch";
  }
  const l = c ? "pre" : "";
  return n.major !== r.major ? l + "major" : n.minor !== r.minor ? l + "minor" : n.patch !== r.patch ? l + "patch" : "prerelease";
};
var SE = AE;
const CE = Fe, bE = (e, t) => new CE(e, t).major;
var NE = bE;
const RE = Fe, $E = (e, t) => new RE(e, t).minor;
var OE = $E;
const IE = Fe, DE = (e, t) => new IE(e, t).patch;
var PE = DE;
const LE = En, FE = (e, t) => {
  const n = LE(e, t);
  return n && n.prerelease.length ? n.prerelease : null;
};
var xE = FE;
const xs = Fe, UE = (e, t, n) => new xs(e, n).compare(new xs(t, n));
var Je = UE;
const kE = Je, ME = (e, t, n) => kE(t, e, n);
var BE = ME;
const HE = Je, jE = (e, t) => HE(e, t, !0);
var GE = jE;
const Us = Fe, qE = (e, t, n) => {
  const r = new Us(e, n), i = new Us(t, n);
  return r.compare(i) || r.compareBuild(i);
};
var na = qE;
const WE = na, VE = (e, t) => e.sort((n, r) => WE(n, r, t));
var zE = VE;
const YE = na, XE = (e, t) => e.sort((n, r) => YE(r, n, t));
var KE = XE;
const JE = Je, QE = (e, t, n) => JE(e, t, n) > 0;
var hi = QE;
const ZE = Je, ey = (e, t, n) => ZE(e, t, n) < 0;
var ra = ey;
const ty = Je, ny = (e, t, n) => ty(e, t, n) === 0;
var Gu = ny;
const ry = Je, iy = (e, t, n) => ry(e, t, n) !== 0;
var qu = iy;
const oy = Je, ay = (e, t, n) => oy(e, t, n) >= 0;
var ia = ay;
const sy = Je, ly = (e, t, n) => sy(e, t, n) <= 0;
var oa = ly;
const cy = Gu, uy = qu, fy = hi, dy = ia, hy = ra, py = oa, my = (e, t, n, r) => {
  switch (t) {
    case "===":
      return typeof e == "object" && (e = e.version), typeof n == "object" && (n = n.version), e === n;
    case "!==":
      return typeof e == "object" && (e = e.version), typeof n == "object" && (n = n.version), e !== n;
    case "":
    case "=":
    case "==":
      return cy(e, n, r);
    case "!=":
      return uy(e, n, r);
    case ">":
      return fy(e, n, r);
    case ">=":
      return dy(e, n, r);
    case "<":
      return hy(e, n, r);
    case "<=":
      return py(e, n, r);
    default:
      throw new TypeError(`Invalid operator: ${t}`);
  }
};
var Wu = my;
const gy = Fe, Ey = En, { safeRe: Ir, t: Dr } = sr, yy = (e, t) => {
  if (e instanceof gy)
    return e;
  if (typeof e == "number" && (e = String(e)), typeof e != "string")
    return null;
  t = t || {};
  let n = null;
  if (!t.rtl)
    n = e.match(t.includePrerelease ? Ir[Dr.COERCEFULL] : Ir[Dr.COERCE]);
  else {
    const c = t.includePrerelease ? Ir[Dr.COERCERTLFULL] : Ir[Dr.COERCERTL];
    let p;
    for (; (p = c.exec(e)) && (!n || n.index + n[0].length !== e.length); )
      (!n || p.index + p[0].length !== n.index + n[0].length) && (n = p), c.lastIndex = p.index + p[1].length + p[2].length;
    c.lastIndex = -1;
  }
  if (n === null)
    return null;
  const r = n[2], i = n[3] || "0", o = n[4] || "0", a = t.includePrerelease && n[5] ? `-${n[5]}` : "", s = t.includePrerelease && n[6] ? `+${n[6]}` : "";
  return Ey(`${r}.${i}.${o}${a}${s}`, t);
};
var _y = yy;
class wy {
  constructor() {
    this.max = 1e3, this.map = /* @__PURE__ */ new Map();
  }
  get(t) {
    const n = this.map.get(t);
    if (n !== void 0)
      return this.map.delete(t), this.map.set(t, n), n;
  }
  delete(t) {
    return this.map.delete(t);
  }
  set(t, n) {
    if (!this.delete(t) && n !== void 0) {
      if (this.map.size >= this.max) {
        const i = this.map.keys().next().value;
        this.delete(i);
      }
      this.map.set(t, n);
    }
    return this;
  }
}
var vy = wy, zi, ks;
function Qe() {
  if (ks) return zi;
  ks = 1;
  const e = /\s+/g;
  class t {
    constructor(N, I) {
      if (I = i(I), N instanceof t)
        return N.loose === !!I.loose && N.includePrerelease === !!I.includePrerelease ? N : new t(N.raw, I);
      if (N instanceof o)
        return this.raw = N.value, this.set = [[N]], this.formatted = void 0, this;
      if (this.options = I, this.loose = !!I.loose, this.includePrerelease = !!I.includePrerelease, this.raw = N.trim().replace(e, " "), this.set = this.raw.split("||").map((b) => this.parseRange(b.trim())).filter((b) => b.length), !this.set.length)
        throw new TypeError(`Invalid SemVer Range: ${this.raw}`);
      if (this.set.length > 1) {
        const b = this.set[0];
        if (this.set = this.set.filter((D) => !E(D[0])), this.set.length === 0)
          this.set = [b];
        else if (this.set.length > 1) {
          for (const D of this.set)
            if (D.length === 1 && T(D[0])) {
              this.set = [D];
              break;
            }
        }
      }
      this.formatted = void 0;
    }
    get range() {
      if (this.formatted === void 0) {
        this.formatted = "";
        for (let N = 0; N < this.set.length; N++) {
          N > 0 && (this.formatted += "||");
          const I = this.set[N];
          for (let b = 0; b < I.length; b++)
            b > 0 && (this.formatted += " "), this.formatted += I[b].toString().trim();
        }
      }
      return this.formatted;
    }
    format() {
      return this.range;
    }
    toString() {
      return this.range;
    }
    parseRange(N) {
      const b = ((this.options.includePrerelease && g) | (this.options.loose && _)) + ":" + N, D = r.get(b);
      if (D)
        return D;
      const O = this.options.loose, k = O ? c[p.HYPHENRANGELOOSE] : c[p.HYPHENRANGE];
      N = N.replace(k, Y(this.options.includePrerelease)), a("hyphen replace", N), N = N.replace(c[p.COMPARATORTRIM], l), a("comparator trim", N), N = N.replace(c[p.TILDETRIM], f), a("tilde trim", N), N = N.replace(c[p.CARETTRIM], h), a("caret trim", N);
      let W = N.split(" ").map((M) => R(M, this.options)).join(" ").split(/\s+/).map((M) => j(M, this.options));
      O && (W = W.filter((M) => (a("loose invalid filter", M, this.options), !!M.match(c[p.COMPARATORLOOSE])))), a("range list", W);
      const L = /* @__PURE__ */ new Map(), X = W.map((M) => new o(M, this.options));
      for (const M of X) {
        if (E(M))
          return [M];
        L.set(M.value, M);
      }
      L.size > 1 && L.has("") && L.delete("");
      const ue = [...L.values()];
      return r.set(b, ue), ue;
    }
    intersects(N, I) {
      if (!(N instanceof t))
        throw new TypeError("a Range is required");
      return this.set.some((b) => A(b, I) && N.set.some((D) => A(D, I) && b.every((O) => D.every((k) => O.intersects(k, I)))));
    }
    // if ANY of the sets match ALL of its comparators, then pass
    test(N) {
      if (!N)
        return !1;
      if (typeof N == "string")
        try {
          N = new s(N, this.options);
        } catch {
          return !1;
        }
      for (let I = 0; I < this.set.length; I++)
        if (te(this.set[I], N, this.options))
          return !0;
      return !1;
    }
  }
  zi = t;
  const n = vy, r = new n(), i = ta, o = pi(), a = di, s = Fe, {
    safeRe: c,
    t: p,
    comparatorTrimReplace: l,
    tildeTrimReplace: f,
    caretTrimReplace: h
  } = sr, { FLAG_INCLUDE_PRERELEASE: g, FLAG_LOOSE: _ } = fi, E = ($) => $.value === "<0.0.0-0", T = ($) => $.value === "", A = ($, N) => {
    let I = !0;
    const b = $.slice();
    let D = b.pop();
    for (; I && b.length; )
      I = b.every((O) => D.intersects(O, N)), D = b.pop();
    return I;
  }, R = ($, N) => ($ = $.replace(c[p.BUILD], ""), a("comp", $, N), $ = Q($, N), a("caret", $), $ = B($, N), a("tildes", $), $ = ae($, N), a("xrange", $), $ = y($, N), a("stars", $), $), P = ($) => !$ || $.toLowerCase() === "x" || $ === "*", B = ($, N) => $.trim().split(/\s+/).map((I) => q(I, N)).join(" "), q = ($, N) => {
    const I = N.loose ? c[p.TILDELOOSE] : c[p.TILDE];
    return $.replace(I, (b, D, O, k, W) => {
      a("tilde", $, b, D, O, k, W);
      let L;
      return P(D) ? L = "" : P(O) ? L = `>=${D}.0.0 <${+D + 1}.0.0-0` : P(k) ? L = `>=${D}.${O}.0 <${D}.${+O + 1}.0-0` : W ? (a("replaceTilde pr", W), L = `>=${D}.${O}.${k}-${W} <${D}.${+O + 1}.0-0`) : L = `>=${D}.${O}.${k} <${D}.${+O + 1}.0-0`, a("tilde return", L), L;
    });
  }, Q = ($, N) => $.trim().split(/\s+/).map((I) => Z(I, N)).join(" "), Z = ($, N) => {
    a("caret", $, N);
    const I = N.loose ? c[p.CARETLOOSE] : c[p.CARET], b = N.includePrerelease ? "-0" : "";
    return $.replace(I, (D, O, k, W, L) => {
      a("caret", $, D, O, k, W, L);
      let X;
      return P(O) ? X = "" : P(k) ? X = `>=${O}.0.0${b} <${+O + 1}.0.0-0` : P(W) ? O === "0" ? X = `>=${O}.${k}.0${b} <${O}.${+k + 1}.0-0` : X = `>=${O}.${k}.0${b} <${+O + 1}.0.0-0` : L ? (a("replaceCaret pr", L), O === "0" ? k === "0" ? X = `>=${O}.${k}.${W}-${L} <${O}.${k}.${+W + 1}-0` : X = `>=${O}.${k}.${W}-${L} <${O}.${+k + 1}.0-0` : X = `>=${O}.${k}.${W}-${L} <${+O + 1}.0.0-0`) : (a("no pr"), O === "0" ? k === "0" ? X = `>=${O}.${k}.${W}${b} <${O}.${k}.${+W + 1}-0` : X = `>=${O}.${k}.${W}${b} <${O}.${+k + 1}.0-0` : X = `>=${O}.${k}.${W} <${+O + 1}.0.0-0`), a("caret return", X), X;
    });
  }, ae = ($, N) => (a("replaceXRanges", $, N), $.split(/\s+/).map((I) => U(I, N)).join(" ")), U = ($, N) => {
    $ = $.trim();
    const I = N.loose ? c[p.XRANGELOOSE] : c[p.XRANGE];
    return $.replace(I, (b, D, O, k, W, L) => {
      a("xRange", $, b, D, O, k, W, L);
      const X = P(O), ue = X || P(k), M = ue || P(W), ve = M;
      return D === "=" && ve && (D = ""), L = N.includePrerelease ? "-0" : "", X ? D === ">" || D === "<" ? b = "<0.0.0-0" : b = "*" : D && ve ? (ue && (k = 0), W = 0, D === ">" ? (D = ">=", ue ? (O = +O + 1, k = 0, W = 0) : (k = +k + 1, W = 0)) : D === "<=" && (D = "<", ue ? O = +O + 1 : k = +k + 1), D === "<" && (L = "-0"), b = `${D + O}.${k}.${W}${L}`) : ue ? b = `>=${O}.0.0${L} <${+O + 1}.0.0-0` : M && (b = `>=${O}.${k}.0${L} <${O}.${+k + 1}.0-0`), a("xRange return", b), b;
    });
  }, y = ($, N) => (a("replaceStars", $, N), $.trim().replace(c[p.STAR], "")), j = ($, N) => (a("replaceGTE0", $, N), $.trim().replace(c[N.includePrerelease ? p.GTE0PRE : p.GTE0], "")), Y = ($) => (N, I, b, D, O, k, W, L, X, ue, M, ve) => (P(b) ? I = "" : P(D) ? I = `>=${b}.0.0${$ ? "-0" : ""}` : P(O) ? I = `>=${b}.${D}.0${$ ? "-0" : ""}` : k ? I = `>=${I}` : I = `>=${I}${$ ? "-0" : ""}`, P(X) ? L = "" : P(ue) ? L = `<${+X + 1}.0.0-0` : P(M) ? L = `<${X}.${+ue + 1}.0-0` : ve ? L = `<=${X}.${ue}.${M}-${ve}` : $ ? L = `<${X}.${ue}.${+M + 1}-0` : L = `<=${L}`, `${I} ${L}`.trim()), te = ($, N, I) => {
    for (let b = 0; b < $.length; b++)
      if (!$[b].test(N))
        return !1;
    if (N.prerelease.length && !I.includePrerelease) {
      for (let b = 0; b < $.length; b++)
        if (a($[b].semver), $[b].semver !== o.ANY && $[b].semver.prerelease.length > 0) {
          const D = $[b].semver;
          if (D.major === N.major && D.minor === N.minor && D.patch === N.patch)
            return !0;
        }
      return !1;
    }
    return !0;
  };
  return zi;
}
var Yi, Ms;
function pi() {
  if (Ms) return Yi;
  Ms = 1;
  const e = Symbol("SemVer ANY");
  class t {
    static get ANY() {
      return e;
    }
    constructor(l, f) {
      if (f = n(f), l instanceof t) {
        if (l.loose === !!f.loose)
          return l;
        l = l.value;
      }
      l = l.trim().split(/\s+/).join(" "), a("comparator", l, f), this.options = f, this.loose = !!f.loose, this.parse(l), this.semver === e ? this.value = "" : this.value = this.operator + this.semver.version, a("comp", this);
    }
    parse(l) {
      const f = this.options.loose ? r[i.COMPARATORLOOSE] : r[i.COMPARATOR], h = l.match(f);
      if (!h)
        throw new TypeError(`Invalid comparator: ${l}`);
      this.operator = h[1] !== void 0 ? h[1] : "", this.operator === "=" && (this.operator = ""), h[2] ? this.semver = new s(h[2], this.options.loose) : this.semver = e;
    }
    toString() {
      return this.value;
    }
    test(l) {
      if (a("Comparator.test", l, this.options.loose), this.semver === e || l === e)
        return !0;
      if (typeof l == "string")
        try {
          l = new s(l, this.options);
        } catch {
          return !1;
        }
      return o(l, this.operator, this.semver, this.options);
    }
    intersects(l, f) {
      if (!(l instanceof t))
        throw new TypeError("a Comparator is required");
      return this.operator === "" ? this.value === "" ? !0 : new c(l.value, f).test(this.value) : l.operator === "" ? l.value === "" ? !0 : new c(this.value, f).test(l.semver) : (f = n(f), f.includePrerelease && (this.value === "<0.0.0-0" || l.value === "<0.0.0-0") || !f.includePrerelease && (this.value.startsWith("<0.0.0") || l.value.startsWith("<0.0.0")) ? !1 : !!(this.operator.startsWith(">") && l.operator.startsWith(">") || this.operator.startsWith("<") && l.operator.startsWith("<") || this.semver.version === l.semver.version && this.operator.includes("=") && l.operator.includes("=") || o(this.semver, "<", l.semver, f) && this.operator.startsWith(">") && l.operator.startsWith("<") || o(this.semver, ">", l.semver, f) && this.operator.startsWith("<") && l.operator.startsWith(">")));
    }
  }
  Yi = t;
  const n = ta, { safeRe: r, t: i } = sr, o = Wu, a = di, s = Fe, c = Qe();
  return Yi;
}
const Ty = Qe(), Ay = (e, t, n) => {
  try {
    t = new Ty(t, n);
  } catch {
    return !1;
  }
  return t.test(e);
};
var mi = Ay;
const Sy = Qe(), Cy = (e, t) => new Sy(e, t).set.map((n) => n.map((r) => r.value).join(" ").trim().split(" "));
var by = Cy;
const Ny = Fe, Ry = Qe(), $y = (e, t, n) => {
  let r = null, i = null, o = null;
  try {
    o = new Ry(t, n);
  } catch {
    return null;
  }
  return e.forEach((a) => {
    o.test(a) && (!r || i.compare(a) === -1) && (r = a, i = new Ny(r, n));
  }), r;
};
var Oy = $y;
const Iy = Fe, Dy = Qe(), Py = (e, t, n) => {
  let r = null, i = null, o = null;
  try {
    o = new Dy(t, n);
  } catch {
    return null;
  }
  return e.forEach((a) => {
    o.test(a) && (!r || i.compare(a) === 1) && (r = a, i = new Iy(r, n));
  }), r;
};
var Ly = Py;
const Xi = Fe, Fy = Qe(), Bs = hi, xy = (e, t) => {
  e = new Fy(e, t);
  let n = new Xi("0.0.0");
  if (e.test(n) || (n = new Xi("0.0.0-0"), e.test(n)))
    return n;
  n = null;
  for (let r = 0; r < e.set.length; ++r) {
    const i = e.set[r];
    let o = null;
    i.forEach((a) => {
      const s = new Xi(a.semver.version);
      switch (a.operator) {
        case ">":
          s.prerelease.length === 0 ? s.patch++ : s.prerelease.push(0), s.raw = s.format();
        case "":
        case ">=":
          (!o || Bs(s, o)) && (o = s);
          break;
        case "<":
        case "<=":
          break;
        default:
          throw new Error(`Unexpected operation: ${a.operator}`);
      }
    }), o && (!n || Bs(n, o)) && (n = o);
  }
  return n && e.test(n) ? n : null;
};
var Uy = xy;
const ky = Qe(), My = (e, t) => {
  try {
    return new ky(e, t).range || "*";
  } catch {
    return null;
  }
};
var By = My;
const Hy = Fe, Vu = pi(), { ANY: jy } = Vu, Gy = Qe(), qy = mi, Hs = hi, js = ra, Wy = oa, Vy = ia, zy = (e, t, n, r) => {
  e = new Hy(e, r), t = new Gy(t, r);
  let i, o, a, s, c;
  switch (n) {
    case ">":
      i = Hs, o = Wy, a = js, s = ">", c = ">=";
      break;
    case "<":
      i = js, o = Vy, a = Hs, s = "<", c = "<=";
      break;
    default:
      throw new TypeError('Must provide a hilo val of "<" or ">"');
  }
  if (qy(e, t, r))
    return !1;
  for (let p = 0; p < t.set.length; ++p) {
    const l = t.set[p];
    let f = null, h = null;
    if (l.forEach((g) => {
      g.semver === jy && (g = new Vu(">=0.0.0")), f = f || g, h = h || g, i(g.semver, f.semver, r) ? f = g : a(g.semver, h.semver, r) && (h = g);
    }), f.operator === s || f.operator === c || (!h.operator || h.operator === s) && o(e, h.semver))
      return !1;
    if (h.operator === c && a(e, h.semver))
      return !1;
  }
  return !0;
};
var aa = zy;
const Yy = aa, Xy = (e, t, n) => Yy(e, t, ">", n);
var Ky = Xy;
const Jy = aa, Qy = (e, t, n) => Jy(e, t, "<", n);
var Zy = Qy;
const Gs = Qe(), e_ = (e, t, n) => (e = new Gs(e, n), t = new Gs(t, n), e.intersects(t, n));
var t_ = e_;
const n_ = mi, r_ = Je;
var i_ = (e, t, n) => {
  const r = [];
  let i = null, o = null;
  const a = e.sort((l, f) => r_(l, f, n));
  for (const l of a)
    n_(l, t, n) ? (o = l, i || (i = l)) : (o && r.push([i, o]), o = null, i = null);
  i && r.push([i, null]);
  const s = [];
  for (const [l, f] of r)
    l === f ? s.push(l) : !f && l === a[0] ? s.push("*") : f ? l === a[0] ? s.push(`<=${f}`) : s.push(`${l} - ${f}`) : s.push(`>=${l}`);
  const c = s.join(" || "), p = typeof t.raw == "string" ? t.raw : String(t);
  return c.length < p.length ? c : t;
};
const qs = Qe(), sa = pi(), { ANY: Ki } = sa, Rn = mi, la = Je, o_ = (e, t, n = {}) => {
  if (e === t)
    return !0;
  e = new qs(e, n), t = new qs(t, n);
  let r = !1;
  e: for (const i of e.set) {
    for (const o of t.set) {
      const a = s_(i, o, n);
      if (r = r || a !== null, a)
        continue e;
    }
    if (r)
      return !1;
  }
  return !0;
}, a_ = [new sa(">=0.0.0-0")], Ws = [new sa(">=0.0.0")], s_ = (e, t, n) => {
  if (e === t)
    return !0;
  if (e.length === 1 && e[0].semver === Ki) {
    if (t.length === 1 && t[0].semver === Ki)
      return !0;
    n.includePrerelease ? e = a_ : e = Ws;
  }
  if (t.length === 1 && t[0].semver === Ki) {
    if (n.includePrerelease)
      return !0;
    t = Ws;
  }
  const r = /* @__PURE__ */ new Set();
  let i, o;
  for (const g of e)
    g.operator === ">" || g.operator === ">=" ? i = Vs(i, g, n) : g.operator === "<" || g.operator === "<=" ? o = zs(o, g, n) : r.add(g.semver);
  if (r.size > 1)
    return null;
  let a;
  if (i && o) {
    if (a = la(i.semver, o.semver, n), a > 0)
      return null;
    if (a === 0 && (i.operator !== ">=" || o.operator !== "<="))
      return null;
  }
  for (const g of r) {
    if (i && !Rn(g, String(i), n) || o && !Rn(g, String(o), n))
      return null;
    for (const _ of t)
      if (!Rn(g, String(_), n))
        return !1;
    return !0;
  }
  let s, c, p, l, f = o && !n.includePrerelease && o.semver.prerelease.length ? o.semver : !1, h = i && !n.includePrerelease && i.semver.prerelease.length ? i.semver : !1;
  f && f.prerelease.length === 1 && o.operator === "<" && f.prerelease[0] === 0 && (f = !1);
  for (const g of t) {
    if (l = l || g.operator === ">" || g.operator === ">=", p = p || g.operator === "<" || g.operator === "<=", i) {
      if (h && g.semver.prerelease && g.semver.prerelease.length && g.semver.major === h.major && g.semver.minor === h.minor && g.semver.patch === h.patch && (h = !1), g.operator === ">" || g.operator === ">=") {
        if (s = Vs(i, g, n), s === g && s !== i)
          return !1;
      } else if (i.operator === ">=" && !Rn(i.semver, String(g), n))
        return !1;
    }
    if (o) {
      if (f && g.semver.prerelease && g.semver.prerelease.length && g.semver.major === f.major && g.semver.minor === f.minor && g.semver.patch === f.patch && (f = !1), g.operator === "<" || g.operator === "<=") {
        if (c = zs(o, g, n), c === g && c !== o)
          return !1;
      } else if (o.operator === "<=" && !Rn(o.semver, String(g), n))
        return !1;
    }
    if (!g.operator && (o || i) && a !== 0)
      return !1;
  }
  return !(i && p && !o && a !== 0 || o && l && !i && a !== 0 || h || f);
}, Vs = (e, t, n) => {
  if (!e)
    return t;
  const r = la(e.semver, t.semver, n);
  return r > 0 ? e : r < 0 || t.operator === ">" && e.operator === ">=" ? t : e;
}, zs = (e, t, n) => {
  if (!e)
    return t;
  const r = la(e.semver, t.semver, n);
  return r < 0 ? e : r > 0 || t.operator === "<" && e.operator === "<=" ? t : e;
};
var l_ = o_;
const Ji = sr, Ys = fi, c_ = Fe, Xs = ju, u_ = En, f_ = EE, d_ = wE, h_ = TE, p_ = SE, m_ = NE, g_ = OE, E_ = PE, y_ = xE, __ = Je, w_ = BE, v_ = GE, T_ = na, A_ = zE, S_ = KE, C_ = hi, b_ = ra, N_ = Gu, R_ = qu, $_ = ia, O_ = oa, I_ = Wu, D_ = _y, P_ = pi(), L_ = Qe(), F_ = mi, x_ = by, U_ = Oy, k_ = Ly, M_ = Uy, B_ = By, H_ = aa, j_ = Ky, G_ = Zy, q_ = t_, W_ = i_, V_ = l_;
var zu = {
  parse: u_,
  valid: f_,
  clean: d_,
  inc: h_,
  diff: p_,
  major: m_,
  minor: g_,
  patch: E_,
  prerelease: y_,
  compare: __,
  rcompare: w_,
  compareLoose: v_,
  compareBuild: T_,
  sort: A_,
  rsort: S_,
  gt: C_,
  lt: b_,
  eq: N_,
  neq: R_,
  gte: $_,
  lte: O_,
  cmp: I_,
  coerce: D_,
  Comparator: P_,
  Range: L_,
  satisfies: F_,
  toComparators: x_,
  maxSatisfying: U_,
  minSatisfying: k_,
  minVersion: M_,
  validRange: B_,
  outside: H_,
  gtr: j_,
  ltr: G_,
  intersects: q_,
  simplifyRange: W_,
  subset: V_,
  SemVer: c_,
  re: Ji.re,
  src: Ji.src,
  tokens: Ji.t,
  SEMVER_SPEC_VERSION: Ys.SEMVER_SPEC_VERSION,
  RELEASE_TYPES: Ys.RELEASE_TYPES,
  compareIdentifiers: Xs.compareIdentifiers,
  rcompareIdentifiers: Xs.rcompareIdentifiers
}, lr = {}, Jr = { exports: {} };
Jr.exports;
(function(e, t) {
  var n = 200, r = "__lodash_hash_undefined__", i = 1, o = 2, a = 9007199254740991, s = "[object Arguments]", c = "[object Array]", p = "[object AsyncFunction]", l = "[object Boolean]", f = "[object Date]", h = "[object Error]", g = "[object Function]", _ = "[object GeneratorFunction]", E = "[object Map]", T = "[object Number]", A = "[object Null]", R = "[object Object]", P = "[object Promise]", B = "[object Proxy]", q = "[object RegExp]", Q = "[object Set]", Z = "[object String]", ae = "[object Symbol]", U = "[object Undefined]", y = "[object WeakMap]", j = "[object ArrayBuffer]", Y = "[object DataView]", te = "[object Float32Array]", $ = "[object Float64Array]", N = "[object Int8Array]", I = "[object Int16Array]", b = "[object Int32Array]", D = "[object Uint8Array]", O = "[object Uint8ClampedArray]", k = "[object Uint16Array]", W = "[object Uint32Array]", L = /[\\^$.*+?()[\]{}|]/g, X = /^\[object .+?Constructor\]$/, ue = /^(?:0|[1-9]\d*)$/, M = {};
  M[te] = M[$] = M[N] = M[I] = M[b] = M[D] = M[O] = M[k] = M[W] = !0, M[s] = M[c] = M[j] = M[l] = M[Y] = M[f] = M[h] = M[g] = M[E] = M[T] = M[R] = M[q] = M[Q] = M[Z] = M[y] = !1;
  var ve = typeof Oe == "object" && Oe && Oe.Object === Object && Oe, wn = typeof self == "object" && self && self.Object === Object && self, je = ve || wn || Function("return this")(), vn = t && !t.nodeType && t, Yt = vn && !0 && e && !e.nodeType && e, dr = Yt && Yt.exports === vn, d = dr && ve.process, u = function() {
    try {
      return d && d.binding && d.binding("util");
    } catch {
    }
  }(), S = u && u.isTypedArray;
  function w(m, v) {
    for (var C = -1, F = m == null ? 0 : m.length, J = 0, H = []; ++C < F; ) {
      var oe = m[C];
      v(oe, C, m) && (H[J++] = oe);
    }
    return H;
  }
  function z(m, v) {
    for (var C = -1, F = v.length, J = m.length; ++C < F; )
      m[J + C] = v[C];
    return m;
  }
  function ne(m, v) {
    for (var C = -1, F = m == null ? 0 : m.length; ++C < F; )
      if (v(m[C], C, m))
        return !0;
    return !1;
  }
  function se(m, v) {
    for (var C = -1, F = Array(m); ++C < m; )
      F[C] = v(C);
    return F;
  }
  function Te(m) {
    return function(v) {
      return m(v);
    };
  }
  function Ae(m, v) {
    return m.has(v);
  }
  function Ge(m, v) {
    return m == null ? void 0 : m[v];
  }
  function fe(m) {
    var v = -1, C = Array(m.size);
    return m.forEach(function(F, J) {
      C[++v] = [J, F];
    }), C;
  }
  function qe(m, v) {
    return function(C) {
      return m(v(C));
    };
  }
  function Ci(m) {
    var v = -1, C = Array(m.size);
    return m.forEach(function(F) {
      C[++v] = F;
    }), C;
  }
  var hr = Array.prototype, Tn = Function.prototype, $t = Object.prototype, bi = je["__core-js_shared__"], ga = Tn.toString, et = $t.hasOwnProperty, Ea = function() {
    var m = /[^.]+$/.exec(bi && bi.keys && bi.keys.IE_PROTO || "");
    return m ? "Symbol(src)_1." + m : "";
  }(), ya = $t.toString, cf = RegExp(
    "^" + ga.call(et).replace(L, "\\$&").replace(/hasOwnProperty|(function).*?(?=\\\()| for .+?(?=\\\])/g, "$1.*?") + "$"
  ), _a = dr ? je.Buffer : void 0, pr = je.Symbol, wa = je.Uint8Array, va = $t.propertyIsEnumerable, uf = hr.splice, Ot = pr ? pr.toStringTag : void 0, Ta = Object.getOwnPropertySymbols, ff = _a ? _a.isBuffer : void 0, df = qe(Object.keys, Object), Ni = Xt(je, "DataView"), An = Xt(je, "Map"), Ri = Xt(je, "Promise"), $i = Xt(je, "Set"), Oi = Xt(je, "WeakMap"), Sn = Xt(Object, "create"), hf = Pt(Ni), pf = Pt(An), mf = Pt(Ri), gf = Pt($i), Ef = Pt(Oi), Aa = pr ? pr.prototype : void 0, Ii = Aa ? Aa.valueOf : void 0;
  function It(m) {
    var v = -1, C = m == null ? 0 : m.length;
    for (this.clear(); ++v < C; ) {
      var F = m[v];
      this.set(F[0], F[1]);
    }
  }
  function yf() {
    this.__data__ = Sn ? Sn(null) : {}, this.size = 0;
  }
  function _f(m) {
    var v = this.has(m) && delete this.__data__[m];
    return this.size -= v ? 1 : 0, v;
  }
  function wf(m) {
    var v = this.__data__;
    if (Sn) {
      var C = v[m];
      return C === r ? void 0 : C;
    }
    return et.call(v, m) ? v[m] : void 0;
  }
  function vf(m) {
    var v = this.__data__;
    return Sn ? v[m] !== void 0 : et.call(v, m);
  }
  function Tf(m, v) {
    var C = this.__data__;
    return this.size += this.has(m) ? 0 : 1, C[m] = Sn && v === void 0 ? r : v, this;
  }
  It.prototype.clear = yf, It.prototype.delete = _f, It.prototype.get = wf, It.prototype.has = vf, It.prototype.set = Tf;
  function ot(m) {
    var v = -1, C = m == null ? 0 : m.length;
    for (this.clear(); ++v < C; ) {
      var F = m[v];
      this.set(F[0], F[1]);
    }
  }
  function Af() {
    this.__data__ = [], this.size = 0;
  }
  function Sf(m) {
    var v = this.__data__, C = gr(v, m);
    if (C < 0)
      return !1;
    var F = v.length - 1;
    return C == F ? v.pop() : uf.call(v, C, 1), --this.size, !0;
  }
  function Cf(m) {
    var v = this.__data__, C = gr(v, m);
    return C < 0 ? void 0 : v[C][1];
  }
  function bf(m) {
    return gr(this.__data__, m) > -1;
  }
  function Nf(m, v) {
    var C = this.__data__, F = gr(C, m);
    return F < 0 ? (++this.size, C.push([m, v])) : C[F][1] = v, this;
  }
  ot.prototype.clear = Af, ot.prototype.delete = Sf, ot.prototype.get = Cf, ot.prototype.has = bf, ot.prototype.set = Nf;
  function Dt(m) {
    var v = -1, C = m == null ? 0 : m.length;
    for (this.clear(); ++v < C; ) {
      var F = m[v];
      this.set(F[0], F[1]);
    }
  }
  function Rf() {
    this.size = 0, this.__data__ = {
      hash: new It(),
      map: new (An || ot)(),
      string: new It()
    };
  }
  function $f(m) {
    var v = Er(this, m).delete(m);
    return this.size -= v ? 1 : 0, v;
  }
  function Of(m) {
    return Er(this, m).get(m);
  }
  function If(m) {
    return Er(this, m).has(m);
  }
  function Df(m, v) {
    var C = Er(this, m), F = C.size;
    return C.set(m, v), this.size += C.size == F ? 0 : 1, this;
  }
  Dt.prototype.clear = Rf, Dt.prototype.delete = $f, Dt.prototype.get = Of, Dt.prototype.has = If, Dt.prototype.set = Df;
  function mr(m) {
    var v = -1, C = m == null ? 0 : m.length;
    for (this.__data__ = new Dt(); ++v < C; )
      this.add(m[v]);
  }
  function Pf(m) {
    return this.__data__.set(m, r), this;
  }
  function Lf(m) {
    return this.__data__.has(m);
  }
  mr.prototype.add = mr.prototype.push = Pf, mr.prototype.has = Lf;
  function ft(m) {
    var v = this.__data__ = new ot(m);
    this.size = v.size;
  }
  function Ff() {
    this.__data__ = new ot(), this.size = 0;
  }
  function xf(m) {
    var v = this.__data__, C = v.delete(m);
    return this.size = v.size, C;
  }
  function Uf(m) {
    return this.__data__.get(m);
  }
  function kf(m) {
    return this.__data__.has(m);
  }
  function Mf(m, v) {
    var C = this.__data__;
    if (C instanceof ot) {
      var F = C.__data__;
      if (!An || F.length < n - 1)
        return F.push([m, v]), this.size = ++C.size, this;
      C = this.__data__ = new Dt(F);
    }
    return C.set(m, v), this.size = C.size, this;
  }
  ft.prototype.clear = Ff, ft.prototype.delete = xf, ft.prototype.get = Uf, ft.prototype.has = kf, ft.prototype.set = Mf;
  function Bf(m, v) {
    var C = yr(m), F = !C && td(m), J = !C && !F && Di(m), H = !C && !F && !J && Da(m), oe = C || F || J || H, pe = oe ? se(m.length, String) : [], Ee = pe.length;
    for (var re in m)
      et.call(m, re) && !(oe && // Safari 9 has enumerable `arguments.length` in strict mode.
      (re == "length" || // Node.js 0.10 has enumerable non-index properties on buffers.
      J && (re == "offset" || re == "parent") || // PhantomJS 2 has enumerable non-index properties on typed arrays.
      H && (re == "buffer" || re == "byteLength" || re == "byteOffset") || // Skip index properties.
      Kf(re, Ee))) && pe.push(re);
    return pe;
  }
  function gr(m, v) {
    for (var C = m.length; C--; )
      if (Ra(m[C][0], v))
        return C;
    return -1;
  }
  function Hf(m, v, C) {
    var F = v(m);
    return yr(m) ? F : z(F, C(m));
  }
  function Cn(m) {
    return m == null ? m === void 0 ? U : A : Ot && Ot in Object(m) ? Yf(m) : ed(m);
  }
  function Sa(m) {
    return bn(m) && Cn(m) == s;
  }
  function Ca(m, v, C, F, J) {
    return m === v ? !0 : m == null || v == null || !bn(m) && !bn(v) ? m !== m && v !== v : jf(m, v, C, F, Ca, J);
  }
  function jf(m, v, C, F, J, H) {
    var oe = yr(m), pe = yr(v), Ee = oe ? c : dt(m), re = pe ? c : dt(v);
    Ee = Ee == s ? R : Ee, re = re == s ? R : re;
    var ke = Ee == R, We = re == R, Se = Ee == re;
    if (Se && Di(m)) {
      if (!Di(v))
        return !1;
      oe = !0, ke = !1;
    }
    if (Se && !ke)
      return H || (H = new ft()), oe || Da(m) ? ba(m, v, C, F, J, H) : Vf(m, v, Ee, C, F, J, H);
    if (!(C & i)) {
      var Me = ke && et.call(m, "__wrapped__"), Be = We && et.call(v, "__wrapped__");
      if (Me || Be) {
        var ht = Me ? m.value() : m, at = Be ? v.value() : v;
        return H || (H = new ft()), J(ht, at, C, F, H);
      }
    }
    return Se ? (H || (H = new ft()), zf(m, v, C, F, J, H)) : !1;
  }
  function Gf(m) {
    if (!Ia(m) || Qf(m))
      return !1;
    var v = $a(m) ? cf : X;
    return v.test(Pt(m));
  }
  function qf(m) {
    return bn(m) && Oa(m.length) && !!M[Cn(m)];
  }
  function Wf(m) {
    if (!Zf(m))
      return df(m);
    var v = [];
    for (var C in Object(m))
      et.call(m, C) && C != "constructor" && v.push(C);
    return v;
  }
  function ba(m, v, C, F, J, H) {
    var oe = C & i, pe = m.length, Ee = v.length;
    if (pe != Ee && !(oe && Ee > pe))
      return !1;
    var re = H.get(m);
    if (re && H.get(v))
      return re == v;
    var ke = -1, We = !0, Se = C & o ? new mr() : void 0;
    for (H.set(m, v), H.set(v, m); ++ke < pe; ) {
      var Me = m[ke], Be = v[ke];
      if (F)
        var ht = oe ? F(Be, Me, ke, v, m, H) : F(Me, Be, ke, m, v, H);
      if (ht !== void 0) {
        if (ht)
          continue;
        We = !1;
        break;
      }
      if (Se) {
        if (!ne(v, function(at, Lt) {
          if (!Ae(Se, Lt) && (Me === at || J(Me, at, C, F, H)))
            return Se.push(Lt);
        })) {
          We = !1;
          break;
        }
      } else if (!(Me === Be || J(Me, Be, C, F, H))) {
        We = !1;
        break;
      }
    }
    return H.delete(m), H.delete(v), We;
  }
  function Vf(m, v, C, F, J, H, oe) {
    switch (C) {
      case Y:
        if (m.byteLength != v.byteLength || m.byteOffset != v.byteOffset)
          return !1;
        m = m.buffer, v = v.buffer;
      case j:
        return !(m.byteLength != v.byteLength || !H(new wa(m), new wa(v)));
      case l:
      case f:
      case T:
        return Ra(+m, +v);
      case h:
        return m.name == v.name && m.message == v.message;
      case q:
      case Z:
        return m == v + "";
      case E:
        var pe = fe;
      case Q:
        var Ee = F & i;
        if (pe || (pe = Ci), m.size != v.size && !Ee)
          return !1;
        var re = oe.get(m);
        if (re)
          return re == v;
        F |= o, oe.set(m, v);
        var ke = ba(pe(m), pe(v), F, J, H, oe);
        return oe.delete(m), ke;
      case ae:
        if (Ii)
          return Ii.call(m) == Ii.call(v);
    }
    return !1;
  }
  function zf(m, v, C, F, J, H) {
    var oe = C & i, pe = Na(m), Ee = pe.length, re = Na(v), ke = re.length;
    if (Ee != ke && !oe)
      return !1;
    for (var We = Ee; We--; ) {
      var Se = pe[We];
      if (!(oe ? Se in v : et.call(v, Se)))
        return !1;
    }
    var Me = H.get(m);
    if (Me && H.get(v))
      return Me == v;
    var Be = !0;
    H.set(m, v), H.set(v, m);
    for (var ht = oe; ++We < Ee; ) {
      Se = pe[We];
      var at = m[Se], Lt = v[Se];
      if (F)
        var Pa = oe ? F(Lt, at, Se, v, m, H) : F(at, Lt, Se, m, v, H);
      if (!(Pa === void 0 ? at === Lt || J(at, Lt, C, F, H) : Pa)) {
        Be = !1;
        break;
      }
      ht || (ht = Se == "constructor");
    }
    if (Be && !ht) {
      var _r = m.constructor, wr = v.constructor;
      _r != wr && "constructor" in m && "constructor" in v && !(typeof _r == "function" && _r instanceof _r && typeof wr == "function" && wr instanceof wr) && (Be = !1);
    }
    return H.delete(m), H.delete(v), Be;
  }
  function Na(m) {
    return Hf(m, id, Xf);
  }
  function Er(m, v) {
    var C = m.__data__;
    return Jf(v) ? C[typeof v == "string" ? "string" : "hash"] : C.map;
  }
  function Xt(m, v) {
    var C = Ge(m, v);
    return Gf(C) ? C : void 0;
  }
  function Yf(m) {
    var v = et.call(m, Ot), C = m[Ot];
    try {
      m[Ot] = void 0;
      var F = !0;
    } catch {
    }
    var J = ya.call(m);
    return F && (v ? m[Ot] = C : delete m[Ot]), J;
  }
  var Xf = Ta ? function(m) {
    return m == null ? [] : (m = Object(m), w(Ta(m), function(v) {
      return va.call(m, v);
    }));
  } : od, dt = Cn;
  (Ni && dt(new Ni(new ArrayBuffer(1))) != Y || An && dt(new An()) != E || Ri && dt(Ri.resolve()) != P || $i && dt(new $i()) != Q || Oi && dt(new Oi()) != y) && (dt = function(m) {
    var v = Cn(m), C = v == R ? m.constructor : void 0, F = C ? Pt(C) : "";
    if (F)
      switch (F) {
        case hf:
          return Y;
        case pf:
          return E;
        case mf:
          return P;
        case gf:
          return Q;
        case Ef:
          return y;
      }
    return v;
  });
  function Kf(m, v) {
    return v = v ?? a, !!v && (typeof m == "number" || ue.test(m)) && m > -1 && m % 1 == 0 && m < v;
  }
  function Jf(m) {
    var v = typeof m;
    return v == "string" || v == "number" || v == "symbol" || v == "boolean" ? m !== "__proto__" : m === null;
  }
  function Qf(m) {
    return !!Ea && Ea in m;
  }
  function Zf(m) {
    var v = m && m.constructor, C = typeof v == "function" && v.prototype || $t;
    return m === C;
  }
  function ed(m) {
    return ya.call(m);
  }
  function Pt(m) {
    if (m != null) {
      try {
        return ga.call(m);
      } catch {
      }
      try {
        return m + "";
      } catch {
      }
    }
    return "";
  }
  function Ra(m, v) {
    return m === v || m !== m && v !== v;
  }
  var td = Sa(/* @__PURE__ */ function() {
    return arguments;
  }()) ? Sa : function(m) {
    return bn(m) && et.call(m, "callee") && !va.call(m, "callee");
  }, yr = Array.isArray;
  function nd(m) {
    return m != null && Oa(m.length) && !$a(m);
  }
  var Di = ff || ad;
  function rd(m, v) {
    return Ca(m, v);
  }
  function $a(m) {
    if (!Ia(m))
      return !1;
    var v = Cn(m);
    return v == g || v == _ || v == p || v == B;
  }
  function Oa(m) {
    return typeof m == "number" && m > -1 && m % 1 == 0 && m <= a;
  }
  function Ia(m) {
    var v = typeof m;
    return m != null && (v == "object" || v == "function");
  }
  function bn(m) {
    return m != null && typeof m == "object";
  }
  var Da = S ? Te(S) : qf;
  function id(m) {
    return nd(m) ? Bf(m) : Wf(m);
  }
  function od() {
    return [];
  }
  function ad() {
    return !1;
  }
  e.exports = rd;
})(Jr, Jr.exports);
var z_ = Jr.exports;
Object.defineProperty(lr, "__esModule", { value: !0 });
lr.DownloadedUpdateHelper = void 0;
lr.createTempUpdateFile = Q_;
const Y_ = er, X_ = Ct, Ks = z_, Ft = Nt, Un = ee;
class K_ {
  constructor(t) {
    this.cacheDir = t, this._file = null, this._packageFile = null, this.versionInfo = null, this.fileInfo = null, this._downloadedFileInfo = null;
  }
  get downloadedFileInfo() {
    return this._downloadedFileInfo;
  }
  get file() {
    return this._file;
  }
  get packageFile() {
    return this._packageFile;
  }
  get cacheDirForPendingUpdate() {
    return Un.join(this.cacheDir, "pending");
  }
  async validateDownloadedPath(t, n, r, i) {
    if (this.versionInfo != null && this.file === t && this.fileInfo != null)
      return Ks(this.versionInfo, n) && Ks(this.fileInfo.info, r.info) && await (0, Ft.pathExists)(t) ? t : null;
    const o = await this.getValidCachedUpdateFile(r, i);
    return o === null ? null : (i.info(`Update has already been downloaded to ${t}).`), this._file = o, o);
  }
  async setDownloadedFile(t, n, r, i, o, a) {
    this._file = t, this._packageFile = n, this.versionInfo = r, this.fileInfo = i, this._downloadedFileInfo = {
      fileName: o,
      sha512: i.info.sha512,
      isAdminRightsRequired: i.info.isAdminRightsRequired === !0
    }, a && await (0, Ft.outputJson)(this.getUpdateInfoFile(), this._downloadedFileInfo);
  }
  async clear() {
    this._file = null, this._packageFile = null, this.versionInfo = null, this.fileInfo = null, await this.cleanCacheDirForPendingUpdate();
  }
  async cleanCacheDirForPendingUpdate() {
    try {
      await (0, Ft.emptyDir)(this.cacheDirForPendingUpdate);
    } catch {
    }
  }
  /**
   * Returns "update-info.json" which is created in the update cache directory's "pending" subfolder after the first update is downloaded.  If the update file does not exist then the cache is cleared and recreated.  If the update file exists then its properties are validated.
   * @param fileInfo
   * @param logger
   */
  async getValidCachedUpdateFile(t, n) {
    const r = this.getUpdateInfoFile();
    if (!await (0, Ft.pathExists)(r))
      return null;
    let o;
    try {
      o = await (0, Ft.readJson)(r);
    } catch (p) {
      let l = "No cached update info available";
      return p.code !== "ENOENT" && (await this.cleanCacheDirForPendingUpdate(), l += ` (error on read: ${p.message})`), n.info(l), null;
    }
    if (!((o == null ? void 0 : o.fileName) !== null))
      return n.warn("Cached update info is corrupted: no fileName, directory for cached update will be cleaned"), await this.cleanCacheDirForPendingUpdate(), null;
    if (t.info.sha512 !== o.sha512)
      return n.info(`Cached update sha512 checksum doesn't match the latest available update. New update must be downloaded. Cached: ${o.sha512}, expected: ${t.info.sha512}. Directory for cached update will be cleaned`), await this.cleanCacheDirForPendingUpdate(), null;
    const s = Un.join(this.cacheDirForPendingUpdate, o.fileName);
    if (!await (0, Ft.pathExists)(s))
      return n.info("Cached update file doesn't exist"), null;
    const c = await J_(s);
    return t.info.sha512 !== c ? (n.warn(`Sha512 checksum doesn't match the latest available update. New update must be downloaded. Cached: ${c}, expected: ${t.info.sha512}`), await this.cleanCacheDirForPendingUpdate(), null) : (this._downloadedFileInfo = o, s);
  }
  getUpdateInfoFile() {
    return Un.join(this.cacheDirForPendingUpdate, "update-info.json");
  }
}
lr.DownloadedUpdateHelper = K_;
function J_(e, t = "sha512", n = "base64", r) {
  return new Promise((i, o) => {
    const a = (0, Y_.createHash)(t);
    a.on("error", o).setEncoding(n), (0, X_.createReadStream)(e, {
      ...r,
      highWaterMark: 1024 * 1024
      /* better to use more memory but hash faster */
    }).on("error", o).on("end", () => {
      a.end(), i(a.read());
    }).pipe(a, { end: !1 });
  });
}
async function Q_(e, t, n) {
  let r = 0, i = Un.join(t, e);
  for (let o = 0; o < 3; o++)
    try {
      return await (0, Ft.unlink)(i), i;
    } catch (a) {
      if (a.code === "ENOENT")
        return i;
      n.warn(`Error on remove temp update file: ${a}`), i = Un.join(t, `${r++}-${e}`);
    }
  return i;
}
var gi = {}, ca = {};
Object.defineProperty(ca, "__esModule", { value: !0 });
ca.getAppCacheDir = ew;
const Qi = ee, Z_ = ti;
function ew() {
  const e = (0, Z_.homedir)();
  let t;
  return process.platform === "win32" ? t = process.env.LOCALAPPDATA || Qi.join(e, "AppData", "Local") : process.platform === "darwin" ? t = Qi.join(e, "Library", "Caches") : t = process.env.XDG_CACHE_HOME || Qi.join(e, ".cache"), t;
}
Object.defineProperty(gi, "__esModule", { value: !0 });
gi.ElectronAppAdapter = void 0;
const Js = ee, tw = ca;
class nw {
  constructor(t = Ht.app) {
    this.app = t;
  }
  whenReady() {
    return this.app.whenReady();
  }
  get version() {
    return this.app.getVersion();
  }
  get name() {
    return this.app.getName();
  }
  get isPackaged() {
    return this.app.isPackaged === !0;
  }
  get appUpdateConfigPath() {
    return this.isPackaged ? Js.join(process.resourcesPath, "app-update.yml") : Js.join(this.app.getAppPath(), "dev-app-update.yml");
  }
  get userDataPath() {
    return this.app.getPath("userData");
  }
  get baseCachePath() {
    return (0, tw.getAppCacheDir)();
  }
  quit() {
    this.app.quit();
  }
  relaunch() {
    this.app.relaunch();
  }
  onQuit(t) {
    this.app.once("quit", (n, r) => t(r));
  }
}
gi.ElectronAppAdapter = nw;
var Yu = {};
(function(e) {
  Object.defineProperty(e, "__esModule", { value: !0 }), e.ElectronHttpExecutor = e.NET_SESSION_NAME = void 0, e.getNetSession = n;
  const t = he;
  e.NET_SESSION_NAME = "electron-updater";
  function n() {
    return Ht.session.fromPartition(e.NET_SESSION_NAME, {
      cache: !1
    });
  }
  class r extends t.HttpExecutor {
    constructor(o) {
      super(), this.proxyLoginCallback = o, this.cachedSession = null;
    }
    async download(o, a, s) {
      return await s.cancellationToken.createPromise((c, p, l) => {
        const f = {
          headers: s.headers || void 0,
          redirect: "manual"
        };
        (0, t.configureRequestUrl)(o, f), (0, t.configureRequestOptions)(f), this.doDownload(f, {
          destination: a,
          options: s,
          onCancel: l,
          callback: (h) => {
            h == null ? c(a) : p(h);
          },
          responseHandler: null
        }, 0);
      });
    }
    createRequest(o, a) {
      o.headers && o.headers.Host && (o.host = o.headers.Host, delete o.headers.Host), this.cachedSession == null && (this.cachedSession = n());
      const s = Ht.net.request({
        ...o,
        session: this.cachedSession
      });
      return s.on("response", a), this.proxyLoginCallback != null && s.on("login", this.proxyLoginCallback), s;
    }
    addRedirectHandlers(o, a, s, c, p) {
      o.on("redirect", (l, f, h) => {
        o.abort(), c > this.maxRedirects ? s(this.createMaxRedirectError()) : p(t.HttpExecutor.prepareRedirectUrlOptions(h, a));
      });
    }
  }
  e.ElectronHttpExecutor = r;
})(Yu);
var cr = {}, Ze = {};
Object.defineProperty(Ze, "__esModule", { value: !0 });
Ze.newBaseUrl = rw;
Ze.newUrlFromBase = iw;
Ze.getChannelFilename = ow;
const Xu = bt;
function rw(e) {
  const t = new Xu.URL(e);
  return t.pathname.endsWith("/") || (t.pathname += "/"), t;
}
function iw(e, t, n = !1) {
  const r = new Xu.URL(e, t), i = t.search;
  return i != null && i.length !== 0 ? r.search = i : n && (r.search = `noCache=${Date.now().toString(32)}`), r;
}
function ow(e) {
  return `${e}.yml`;
}
var ce = {}, aw = "[object Symbol]", Ku = /[\\^$.*+?()[\]{}|]/g, sw = RegExp(Ku.source), lw = typeof Oe == "object" && Oe && Oe.Object === Object && Oe, cw = typeof self == "object" && self && self.Object === Object && self, uw = lw || cw || Function("return this")(), fw = Object.prototype, dw = fw.toString, Qs = uw.Symbol, Zs = Qs ? Qs.prototype : void 0, el = Zs ? Zs.toString : void 0;
function hw(e) {
  if (typeof e == "string")
    return e;
  if (mw(e))
    return el ? el.call(e) : "";
  var t = e + "";
  return t == "0" && 1 / e == -1 / 0 ? "-0" : t;
}
function pw(e) {
  return !!e && typeof e == "object";
}
function mw(e) {
  return typeof e == "symbol" || pw(e) && dw.call(e) == aw;
}
function gw(e) {
  return e == null ? "" : hw(e);
}
function Ew(e) {
  return e = gw(e), e && sw.test(e) ? e.replace(Ku, "\\$&") : e;
}
var Ju = Ew;
Object.defineProperty(ce, "__esModule", { value: !0 });
ce.Provider = void 0;
ce.findFile = Tw;
ce.parseUpdateInfo = Aw;
ce.getFileList = Qu;
ce.resolveFiles = Sw;
const At = he, yw = we, _w = bt, Qr = Ze, ww = Ju;
class vw {
  constructor(t) {
    this.runtimeOptions = t, this.requestHeaders = null, this.executor = t.executor;
  }
  // By default, the blockmap file is in the same directory as the main file
  // But some providers may have a different blockmap file, so we need to override this method
  getBlockMapFiles(t, n, r, i = null) {
    const o = (0, Qr.newUrlFromBase)(`${t.pathname}.blockmap`, t);
    return [(0, Qr.newUrlFromBase)(`${t.pathname.replace(new RegExp(ww(r), "g"), n)}.blockmap`, i ? new _w.URL(i) : t), o];
  }
  get isUseMultipleRangeRequest() {
    return this.runtimeOptions.isUseMultipleRangeRequest !== !1;
  }
  getChannelFilePrefix() {
    if (this.runtimeOptions.platform === "linux") {
      const t = process.env.TEST_UPDATER_ARCH || process.arch;
      return "-linux" + (t === "x64" ? "" : `-${t}`);
    } else
      return this.runtimeOptions.platform === "darwin" ? "-mac" : "";
  }
  // due to historical reasons for windows we use channel name without platform specifier
  getDefaultChannelName() {
    return this.getCustomChannelName("latest");
  }
  getCustomChannelName(t) {
    return `${t}${this.getChannelFilePrefix()}`;
  }
  get fileExtraDownloadHeaders() {
    return null;
  }
  setRequestHeaders(t) {
    this.requestHeaders = t;
  }
  /**
   * Method to perform API request only to resolve update info, but not to download update.
   */
  httpRequest(t, n, r) {
    return this.executor.request(this.createRequestOptions(t, n), r);
  }
  createRequestOptions(t, n) {
    const r = {};
    return this.requestHeaders == null ? n != null && (r.headers = n) : r.headers = n == null ? this.requestHeaders : { ...this.requestHeaders, ...n }, (0, At.configureRequestUrl)(t, r), r;
  }
}
ce.Provider = vw;
function Tw(e, t, n) {
  var r;
  if (e.length === 0)
    throw (0, At.newError)("No files provided", "ERR_UPDATER_NO_FILES_PROVIDED");
  const i = e.filter((a) => a.url.pathname.toLowerCase().endsWith(`.${t.toLowerCase()}`)), o = (r = i.find((a) => [a.url.pathname, a.info.url].some((s) => s.includes(process.arch)))) !== null && r !== void 0 ? r : i.shift();
  return o || (n == null ? e[0] : e.find((a) => !n.some((s) => a.url.pathname.toLowerCase().endsWith(`.${s.toLowerCase()}`))));
}
function Aw(e, t, n) {
  if (e == null)
    throw (0, At.newError)(`Cannot parse update info from ${t} in the latest release artifacts (${n}): rawData: null`, "ERR_UPDATER_INVALID_UPDATE_INFO");
  let r;
  try {
    r = (0, yw.load)(e);
  } catch (i) {
    throw (0, At.newError)(`Cannot parse update info from ${t} in the latest release artifacts (${n}): ${i.stack || i.message}, rawData: ${e}`, "ERR_UPDATER_INVALID_UPDATE_INFO");
  }
  return r;
}
function Qu(e) {
  const t = e.files;
  if (t != null && t.length > 0)
    return t;
  if (e.path != null)
    return [
      {
        url: e.path,
        sha2: e.sha2,
        sha512: e.sha512
      }
    ];
  throw (0, At.newError)(`No files provided: ${(0, At.safeStringifyJson)(e)}`, "ERR_UPDATER_NO_FILES_PROVIDED");
}
function Sw(e, t, n = (r) => r) {
  const i = Qu(e).map((s) => {
    if (s.sha2 == null && s.sha512 == null)
      throw (0, At.newError)(`Update info doesn't contain nor sha256 neither sha512 checksum: ${(0, At.safeStringifyJson)(s)}`, "ERR_UPDATER_NO_CHECKSUM");
    return {
      url: (0, Qr.newUrlFromBase)(n(s.url), t),
      info: s
    };
  }), o = e.packages, a = o == null ? null : o[process.arch] || o.ia32;
  return a != null && (i[0].packageInfo = {
    ...a,
    path: (0, Qr.newUrlFromBase)(n(a.path), t).href
  }), i;
}
Object.defineProperty(cr, "__esModule", { value: !0 });
cr.GenericProvider = void 0;
const tl = he, Zi = Ze, eo = ce;
class Cw extends eo.Provider {
  constructor(t, n, r) {
    super(r), this.configuration = t, this.updater = n, this.baseUrl = (0, Zi.newBaseUrl)(this.configuration.url);
  }
  get channel() {
    const t = this.updater.channel || this.configuration.channel;
    return t == null ? this.getDefaultChannelName() : this.getCustomChannelName(t);
  }
  async getLatestVersion() {
    const t = (0, Zi.getChannelFilename)(this.channel), n = (0, Zi.newUrlFromBase)(t, this.baseUrl, this.updater.isAddNoCacheQuery);
    for (let r = 0; ; r++)
      try {
        return (0, eo.parseUpdateInfo)(await this.httpRequest(n), t, n);
      } catch (i) {
        if (i instanceof tl.HttpError && i.statusCode === 404)
          throw (0, tl.newError)(`Cannot find channel "${t}" update info: ${i.stack || i.message}`, "ERR_UPDATER_CHANNEL_FILE_NOT_FOUND");
        if (i.code === "ECONNREFUSED" && r < 3) {
          await new Promise((o, a) => {
            try {
              setTimeout(o, 1e3 * r);
            } catch (s) {
              a(s);
            }
          });
          continue;
        }
        throw i;
      }
  }
  resolveFiles(t) {
    return (0, eo.resolveFiles)(t, this.baseUrl);
  }
}
cr.GenericProvider = Cw;
var Ei = {}, yi = {};
Object.defineProperty(yi, "__esModule", { value: !0 });
yi.BitbucketProvider = void 0;
const nl = he, to = Ze, no = ce;
class bw extends no.Provider {
  constructor(t, n, r) {
    super({
      ...r,
      isUseMultipleRangeRequest: !1
    }), this.configuration = t, this.updater = n;
    const { owner: i, slug: o } = t;
    this.baseUrl = (0, to.newBaseUrl)(`https://api.bitbucket.org/2.0/repositories/${i}/${o}/downloads`);
  }
  get channel() {
    return this.updater.channel || this.configuration.channel || "latest";
  }
  async getLatestVersion() {
    const t = new nl.CancellationToken(), n = (0, to.getChannelFilename)(this.getCustomChannelName(this.channel)), r = (0, to.newUrlFromBase)(n, this.baseUrl, this.updater.isAddNoCacheQuery);
    try {
      const i = await this.httpRequest(r, void 0, t);
      return (0, no.parseUpdateInfo)(i, n, r);
    } catch (i) {
      throw (0, nl.newError)(`Unable to find latest version on ${this.toString()}, please ensure release exists: ${i.stack || i.message}`, "ERR_UPDATER_LATEST_VERSION_NOT_FOUND");
    }
  }
  resolveFiles(t) {
    return (0, no.resolveFiles)(t, this.baseUrl);
  }
  toString() {
    const { owner: t, slug: n } = this.configuration;
    return `Bitbucket (owner: ${t}, slug: ${n}, channel: ${this.channel})`;
  }
}
yi.BitbucketProvider = bw;
var St = {};
Object.defineProperty(St, "__esModule", { value: !0 });
St.GitHubProvider = St.BaseGitHubProvider = void 0;
St.computeReleaseNotes = ef;
const st = he, nt = zu, Nw = bt, sn = Ze, Oo = ce, ro = /\/tag\/(v?[^/]+)$/;
class Zu extends Oo.Provider {
  constructor(t, n, r) {
    super({
      ...r,
      /* because GitHib uses S3 */
      isUseMultipleRangeRequest: !1
    }), this.options = t, this.baseUrl = (0, sn.newBaseUrl)((0, st.githubUrl)(t, n));
    const i = n === "github.com" ? "api.github.com" : n;
    this.baseApiUrl = (0, sn.newBaseUrl)((0, st.githubUrl)(t, i));
  }
  computeGithubBasePath(t) {
    const n = this.options.host;
    return n && !["github.com", "api.github.com"].includes(n) ? `/api/v3${t}` : t;
  }
}
St.BaseGitHubProvider = Zu;
class Rw extends Zu {
  constructor(t, n, r) {
    super(t, "github.com", r), this.options = t, this.updater = n;
  }
  get channel() {
    const t = this.updater.channel || this.options.channel;
    return t == null ? this.getDefaultChannelName() : this.getCustomChannelName(t);
  }
  async getLatestVersion() {
    var t, n, r, i, o;
    const a = new st.CancellationToken(), s = await this.httpRequest((0, sn.newUrlFromBase)(`${this.basePath}.atom`, this.baseUrl), {
      accept: "application/xml, application/atom+xml, text/xml, */*"
    }, a), c = (0, st.parseXml)(s);
    let p = c.element("entry", !1, "No published versions on GitHub"), l = null;
    try {
      if (this.updater.allowPrerelease) {
        const T = ((t = this.updater) === null || t === void 0 ? void 0 : t.channel) || ((n = nt.prerelease(this.updater.currentVersion)) === null || n === void 0 ? void 0 : n[0]) || null;
        if (T === null)
          l = ro.exec(p.element("link").attribute("href"))[1];
        else
          for (const A of c.getElements("entry")) {
            const R = ro.exec(A.element("link").attribute("href"));
            if (R === null)
              continue;
            const P = R[1];
            if (!nt.valid(P))
              continue;
            const B = ((r = nt.prerelease(P)) === null || r === void 0 ? void 0 : r[0]) || null, q = !T || ["alpha", "beta"].includes(T), Q = B !== null && !["alpha", "beta"].includes(String(B));
            if (q && !Q && !(T === "beta" && B === "alpha")) {
              l = P, p = A;
              break;
            }
            if (B && B === T) {
              l = P, p = A;
              break;
            }
          }
      } else {
        l = await this.getLatestTagName(a);
        for (const T of c.getElements("entry")) {
          const A = ro.exec(T.element("link").attribute("href"));
          if (A != null && A[1] === l) {
            p = T;
            break;
          }
        }
      }
    } catch (T) {
      throw (0, st.newError)(`Cannot parse releases feed: ${T.stack || T.message},
XML:
${s}`, "ERR_UPDATER_INVALID_RELEASE_FEED");
    }
    if (l == null)
      throw (0, st.newError)("No published versions on GitHub", "ERR_UPDATER_NO_PUBLISHED_VERSIONS");
    let f, h = "", g = "";
    const _ = async (T) => {
      h = (0, sn.getChannelFilename)(T), g = (0, sn.newUrlFromBase)(this.getBaseDownloadPath(String(l), h), this.baseUrl);
      const A = this.createRequestOptions(g);
      try {
        return await this.executor.request(A, a);
      } catch (R) {
        throw R instanceof st.HttpError && R.statusCode === 404 ? (0, st.newError)(`Cannot find ${h} in the latest release artifacts (${g}): ${R.stack || R.message}`, "ERR_UPDATER_CHANNEL_FILE_NOT_FOUND") : R;
      }
    };
    try {
      let T = this.channel;
      this.updater.allowPrerelease && (!((i = nt.prerelease(l)) === null || i === void 0) && i[0]) && (T = this.getCustomChannelName(String((o = nt.prerelease(l)) === null || o === void 0 ? void 0 : o[0]))), f = await _(T);
    } catch (T) {
      if (this.updater.allowPrerelease)
        f = await _(this.getDefaultChannelName());
      else
        throw T;
    }
    const E = (0, Oo.parseUpdateInfo)(f, h, g);
    return E.releaseName == null && (E.releaseName = p.elementValueOrEmpty("title")), E.releaseNotes == null && (E.releaseNotes = ef(this.updater.currentVersion, this.updater.fullChangelog, c, p)), {
      tag: l,
      ...E
    };
  }
  async getLatestTagName(t) {
    const n = this.options, r = n.host == null || n.host === "github.com" ? (0, sn.newUrlFromBase)(`${this.basePath}/latest`, this.baseUrl) : new Nw.URL(`${this.computeGithubBasePath(`/repos/${n.owner}/${n.repo}/releases`)}/latest`, this.baseApiUrl);
    try {
      const i = await this.httpRequest(r, { Accept: "application/json" }, t);
      return i == null ? null : JSON.parse(i).tag_name;
    } catch (i) {
      throw (0, st.newError)(`Unable to find latest version on GitHub (${r}), please ensure a production release exists: ${i.stack || i.message}`, "ERR_UPDATER_LATEST_VERSION_NOT_FOUND");
    }
  }
  get basePath() {
    return `/${this.options.owner}/${this.options.repo}/releases`;
  }
  resolveFiles(t) {
    return (0, Oo.resolveFiles)(t, this.baseUrl, (n) => this.getBaseDownloadPath(t.tag, n.replace(/ /g, "-")));
  }
  getBaseDownloadPath(t, n) {
    return `${this.basePath}/download/${t}/${n}`;
  }
}
St.GitHubProvider = Rw;
function rl(e) {
  const t = e.elementValueOrEmpty("content");
  return t === "No content." ? "" : t;
}
function ef(e, t, n, r) {
  if (!t)
    return rl(r);
  const i = /\/tag\/v?([^/]+)$/;
  let o;
  try {
    o = i.exec(r.element("link").attribute("href"))[1], o = nt.valid(o) ? o : void 0;
  } catch {
  }
  if (o == null)
    return null;
  const a = [];
  for (const s of n.getElements("entry")) {
    let c;
    try {
      const f = i.exec(s.element("link").attribute("href"));
      if (!f)
        continue;
      c = f[1];
    } catch {
      continue;
    }
    if (!nt.valid(c))
      continue;
    const p = nt.gt(c, e.raw), l = nt.lte(c, o);
    p && l && a.push({
      version: c,
      note: rl(s)
    });
  }
  return a.sort((s, c) => nt.rcompare(s.version, c.version));
}
var _i = {};
Object.defineProperty(_i, "__esModule", { value: !0 });
_i.GitLabProvider = void 0;
const ye = he, io = bt, $w = Ju, Pr = Ze, oo = ce;
class Ow extends oo.Provider {
  /**
   * Normalizes filenames by replacing spaces and underscores with dashes.
   *
   * This is a workaround to handle filename formatting differences between tools:
   * - electron-builder formats filenames like "test file.txt" as "test-file.txt"
   * - GitLab may provide asset URLs using underscores, such as "test_file.txt"
   *
   * Because of this mismatch, we can't reliably extract the correct filename from
   * the asset path without normalization. This function ensures consistent matching
   * across different filename formats by converting all spaces and underscores to dashes.
   *
   * @param filename The filename to normalize
   * @returns The normalized filename with spaces and underscores replaced by dashes
   */
  normalizeFilename(t) {
    return t.replace(/ |_/g, "-");
  }
  constructor(t, n, r) {
    super({
      ...r,
      // GitLab might not support multiple range requests efficiently
      isUseMultipleRangeRequest: !1
    }), this.options = t, this.updater = n, this.cachedLatestVersion = null;
    const o = t.host || "gitlab.com";
    this.baseApiUrl = (0, Pr.newBaseUrl)(`https://${o}/api/v4`);
  }
  createRequestOptions(t, n) {
    const r = super.createRequestOptions(t, n);
    return r.redirect = "manual", r;
  }
  get channel() {
    const t = this.updater.channel || this.options.channel;
    return t == null ? this.getDefaultChannelName() : this.getCustomChannelName(t);
  }
  async getLatestVersion() {
    const t = new ye.CancellationToken(), n = (0, Pr.newUrlFromBase)(`projects/${this.options.projectId}/releases/permalink/latest`, this.baseApiUrl), r = { Accept: "application/json", ...this.setAuthHeaderForToken(this.options.token || null) };
    let i;
    try {
      i = await this.httpRequest(n, r, t);
    } catch (g) {
      throw (0, ye.newError)(`Unable to find latest release on GitLab (${n}): ${g.stack || g.message}`, "ERR_UPDATER_LATEST_VERSION_NOT_FOUND");
    }
    if (!i)
      throw (0, ye.newError)("No published releases on GitLab", "ERR_UPDATER_NO_PUBLISHED_VERSIONS");
    let o;
    try {
      o = JSON.parse(i);
    } catch (g) {
      throw (0, ye.newError)(`Unable to parse latest release response from GitLab (${n}): response was not valid JSON: ${g.stack || g.message}`, "ERR_UPDATER_LATEST_VERSION_NOT_FOUND");
    }
    if (o.upcoming_release)
      throw (0, ye.newError)("Latest GitLab release is scheduled but not yet published", "ERR_UPDATER_LATEST_VERSION_NOT_FOUND");
    const a = o.tag_name;
    let s = null, c = "", p = null;
    const l = async (g) => {
      c = (0, Pr.getChannelFilename)(g);
      const _ = o.assets.links.find((A) => A.name === c);
      if (!_)
        throw (0, ye.newError)(`Cannot find ${c} in the latest release assets`, "ERR_UPDATER_CHANNEL_FILE_NOT_FOUND");
      p = new io.URL(_.direct_asset_url);
      const E = this.setAuthHeaderForToken(this.options.token || null), T = Object.keys(E).length ? E : void 0;
      try {
        const A = await this.httpRequest(p, T, t);
        if (!A)
          throw (0, ye.newError)(`Empty response from ${p}`, "ERR_UPDATER_CHANNEL_FILE_NOT_FOUND");
        return A;
      } catch (A) {
        throw A instanceof ye.HttpError && A.statusCode === 404 ? (0, ye.newError)(`Cannot find ${c} in the latest release artifacts (${p}): ${A.stack || A.message}`, "ERR_UPDATER_CHANNEL_FILE_NOT_FOUND") : A;
      }
    };
    try {
      s = await l(this.channel);
    } catch (g) {
      if (this.channel !== this.getDefaultChannelName())
        s = await l(this.getDefaultChannelName());
      else
        throw g;
    }
    if (!s)
      throw (0, ye.newError)(`Unable to parse channel data from ${c}`, "ERR_UPDATER_INVALID_UPDATE_INFO");
    const f = (0, oo.parseUpdateInfo)(s, c, p);
    f.releaseName == null && (f.releaseName = o.name), f.releaseNotes == null && (f.releaseNotes = o.description || null);
    const h = {
      tag: a,
      assets: this.convertAssetsToMap(o.assets),
      ...f
    };
    return this.cachedLatestVersion = h, h;
  }
  /**
   * Utility function to convert GitlabReleaseAsset to Map<string, string>
   * Maps asset names to their download URLs
   */
  convertAssetsToMap(t) {
    const n = /* @__PURE__ */ new Map();
    for (const r of t.links)
      n.set(this.normalizeFilename(r.name), r.direct_asset_url);
    return n;
  }
  /**
   * Find blockmap file URL in assets map for a specific filename
   */
  findBlockMapInAssets(t, n) {
    const r = [`${n}.blockmap`, `${this.normalizeFilename(n)}.blockmap`];
    for (const i of r) {
      const o = t.get(i);
      if (o)
        return new io.URL(o);
    }
    return null;
  }
  async fetchReleaseInfoByVersion(t) {
    const n = new ye.CancellationToken(), r = [`v${t}`, t];
    for (const i of r) {
      const o = (0, Pr.newUrlFromBase)(`projects/${this.options.projectId}/releases/${encodeURIComponent(i)}`, this.baseApiUrl);
      try {
        const a = { Accept: "application/json", ...this.setAuthHeaderForToken(this.options.token || null) }, s = await this.httpRequest(o, a, n);
        if (s)
          return JSON.parse(s);
      } catch (a) {
        if (a instanceof ye.HttpError && a.statusCode === 404)
          continue;
        throw (0, ye.newError)(`Unable to find release ${i} on GitLab (${o}): ${a.stack || a.message}`, "ERR_UPDATER_RELEASE_NOT_FOUND");
      }
    }
    throw (0, ye.newError)(`Unable to find release with version ${t} (tried: ${r.join(", ")}) on GitLab`, "ERR_UPDATER_RELEASE_NOT_FOUND");
  }
  setAuthHeaderForToken(t) {
    const n = {};
    return t != null && (t.startsWith("Bearer") ? n.authorization = t : n["PRIVATE-TOKEN"] = t), n;
  }
  /**
   * Get version info for blockmap files, using cache when possible
   */
  async getVersionInfoForBlockMap(t) {
    if (this.cachedLatestVersion && this.cachedLatestVersion.version === t)
      return this.cachedLatestVersion.assets;
    const n = await this.fetchReleaseInfoByVersion(t);
    return n && n.assets ? this.convertAssetsToMap(n.assets) : null;
  }
  /**
   * Find blockmap URLs from version assets
   */
  async findBlockMapUrlsFromAssets(t, n, r) {
    let i = null, o = null;
    const a = await this.getVersionInfoForBlockMap(n);
    a && (i = this.findBlockMapInAssets(a, r));
    const s = await this.getVersionInfoForBlockMap(t);
    if (s) {
      const c = r.replace(new RegExp($w(n), "g"), t);
      o = this.findBlockMapInAssets(s, c);
    }
    return [o, i];
  }
  async getBlockMapFiles(t, n, r, i = null) {
    if (this.options.uploadTarget === "project_upload") {
      const o = t.pathname.split("/").pop() || "", [a, s] = await this.findBlockMapUrlsFromAssets(n, r, o);
      if (!s)
        throw (0, ye.newError)(`Cannot find blockmap file for ${r} in GitLab assets`, "ERR_UPDATER_BLOCKMAP_FILE_NOT_FOUND");
      if (!a)
        throw (0, ye.newError)(`Cannot find blockmap file for ${n} in GitLab assets`, "ERR_UPDATER_BLOCKMAP_FILE_NOT_FOUND");
      return [a, s];
    } else
      return super.getBlockMapFiles(t, n, r, i);
  }
  resolveFiles(t) {
    return (0, oo.getFileList)(t).map((n) => {
      const i = [
        n.url,
        // Original filename
        this.normalizeFilename(n.url)
        // Normalized filename (spaces/underscores → dashes)
      ].find((a) => t.assets.has(a)), o = i ? t.assets.get(i) : void 0;
      if (!o)
        throw (0, ye.newError)(`Cannot find asset "${n.url}" in GitLab release assets. Available assets: ${Array.from(t.assets.keys()).join(", ")}`, "ERR_UPDATER_ASSET_NOT_FOUND");
      return {
        url: new io.URL(o),
        info: n
      };
    });
  }
  toString() {
    return `GitLab (projectId: ${this.options.projectId}, channel: ${this.channel})`;
  }
}
_i.GitLabProvider = Ow;
var wi = {};
Object.defineProperty(wi, "__esModule", { value: !0 });
wi.KeygenProvider = void 0;
const il = he, ao = Ze, so = ce;
class Iw extends so.Provider {
  constructor(t, n, r) {
    super({
      ...r,
      isUseMultipleRangeRequest: !1
    }), this.configuration = t, this.updater = n, this.defaultHostname = "api.keygen.sh";
    const i = this.configuration.host || this.defaultHostname;
    this.baseUrl = (0, ao.newBaseUrl)(`https://${i}/v1/accounts/${this.configuration.account}/artifacts?product=${this.configuration.product}`);
  }
  get channel() {
    return this.updater.channel || this.configuration.channel || "stable";
  }
  async getLatestVersion() {
    const t = new il.CancellationToken(), n = (0, ao.getChannelFilename)(this.getCustomChannelName(this.channel)), r = (0, ao.newUrlFromBase)(n, this.baseUrl, this.updater.isAddNoCacheQuery);
    try {
      const i = await this.httpRequest(r, {
        Accept: "application/vnd.api+json",
        "Keygen-Version": "1.1"
      }, t);
      return (0, so.parseUpdateInfo)(i, n, r);
    } catch (i) {
      throw (0, il.newError)(`Unable to find latest version on ${this.toString()}, please ensure release exists: ${i.stack || i.message}`, "ERR_UPDATER_LATEST_VERSION_NOT_FOUND");
    }
  }
  resolveFiles(t) {
    return (0, so.resolveFiles)(t, this.baseUrl);
  }
  toString() {
    const { account: t, product: n, platform: r } = this.configuration;
    return `Keygen (account: ${t}, product: ${n}, platform: ${r}, channel: ${this.channel})`;
  }
}
wi.KeygenProvider = Iw;
var vi = {};
Object.defineProperty(vi, "__esModule", { value: !0 });
vi.PrivateGitHubProvider = void 0;
const Jt = he, Dw = we, Pw = ee, ol = bt, al = Ze, Lw = St, Fw = ce;
class xw extends Lw.BaseGitHubProvider {
  constructor(t, n, r, i) {
    super(t, "api.github.com", i), this.updater = n, this.token = r;
  }
  createRequestOptions(t, n) {
    const r = super.createRequestOptions(t, n);
    return r.redirect = "manual", r;
  }
  async getLatestVersion() {
    const t = new Jt.CancellationToken(), n = (0, al.getChannelFilename)(this.getDefaultChannelName()), r = await this.getLatestVersionInfo(t), i = r.assets.find((s) => s.name === n);
    if (i == null)
      throw (0, Jt.newError)(`Cannot find ${n} in the release ${r.html_url || r.name}`, "ERR_UPDATER_CHANNEL_FILE_NOT_FOUND");
    const o = new ol.URL(i.url);
    let a;
    try {
      a = (0, Dw.load)(await this.httpRequest(o, this.configureHeaders("application/octet-stream"), t));
    } catch (s) {
      throw s instanceof Jt.HttpError && s.statusCode === 404 ? (0, Jt.newError)(`Cannot find ${n} in the latest release artifacts (${o}): ${s.stack || s.message}`, "ERR_UPDATER_CHANNEL_FILE_NOT_FOUND") : s;
    }
    return a.assets = r.assets, a;
  }
  get fileExtraDownloadHeaders() {
    return this.configureHeaders("application/octet-stream");
  }
  configureHeaders(t) {
    return {
      accept: t,
      authorization: `token ${this.token}`
    };
  }
  async getLatestVersionInfo(t) {
    const n = this.updater.allowPrerelease;
    let r = this.basePath;
    n || (r = `${r}/latest`);
    const i = (0, al.newUrlFromBase)(r, this.baseUrl);
    try {
      const o = JSON.parse(await this.httpRequest(i, this.configureHeaders("application/vnd.github.v3+json"), t));
      if (n) {
        const a = o.filter((s) => !s.draft);
        return a.find((s) => s.prerelease) || a[0];
      } else
        return o;
    } catch (o) {
      throw (0, Jt.newError)(`Unable to find latest version on GitHub (${i}), please ensure a production release exists: ${o.stack || o.message}`, "ERR_UPDATER_LATEST_VERSION_NOT_FOUND");
    }
  }
  get basePath() {
    return this.computeGithubBasePath(`/repos/${this.options.owner}/${this.options.repo}/releases`);
  }
  resolveFiles(t) {
    return (0, Fw.getFileList)(t).map((n) => {
      const r = Pw.posix.basename(n.url).replace(/ /g, "-"), i = t.assets.find((o) => o != null && o.name === r);
      if (i == null)
        throw (0, Jt.newError)(`Cannot find asset "${r}" in: ${JSON.stringify(t.assets, null, 2)}`, "ERR_UPDATER_ASSET_NOT_FOUND");
      return {
        url: new ol.URL(i.url),
        info: n
      };
    });
  }
}
vi.PrivateGitHubProvider = xw;
Object.defineProperty(Ei, "__esModule", { value: !0 });
Ei.isUrlProbablySupportMultiRangeRequests = tf;
Ei.createClient = jw;
const Lr = he, Uw = yi, sl = cr, kw = St, Mw = _i, Bw = wi, Hw = vi;
function tf(e) {
  return !e.includes("s3.amazonaws.com");
}
function jw(e, t, n) {
  if (typeof e == "string")
    throw (0, Lr.newError)("Please pass PublishConfiguration object", "ERR_UPDATER_INVALID_PROVIDER_CONFIGURATION");
  const r = e.provider;
  switch (r) {
    case "github": {
      const i = e, o = (i.private ? process.env.GH_TOKEN || process.env.GITHUB_TOKEN : null) || i.token;
      return o == null ? new kw.GitHubProvider(i, t, n) : new Hw.PrivateGitHubProvider(i, t, o, n);
    }
    case "bitbucket":
      return new Uw.BitbucketProvider(e, t, n);
    case "gitlab":
      return new Mw.GitLabProvider(e, t, n);
    case "keygen":
      return new Bw.KeygenProvider(e, t, n);
    case "s3":
    case "spaces":
      return new sl.GenericProvider({
        provider: "generic",
        url: (0, Lr.getS3LikeProviderBaseUrl)(e),
        channel: e.channel || null
      }, t, {
        ...n,
        // https://github.com/minio/minio/issues/5285#issuecomment-350428955
        isUseMultipleRangeRequest: !1
      });
    case "generic": {
      const i = e;
      return new sl.GenericProvider(i, t, {
        ...n,
        isUseMultipleRangeRequest: i.useMultipleRangeRequest !== !1 && tf(i.url)
      });
    }
    case "custom": {
      const i = e, o = i.updateProvider;
      if (!o)
        throw (0, Lr.newError)("Custom provider not specified", "ERR_UPDATER_INVALID_PROVIDER_CONFIGURATION");
      return new o(i, t, n);
    }
    default:
      throw (0, Lr.newError)(`Unsupported provider: ${r}`, "ERR_UPDATER_UNSUPPORTED_PROVIDER");
  }
}
var Ti = {}, ur = {}, yn = {}, zt = {};
Object.defineProperty(zt, "__esModule", { value: !0 });
zt.OperationKind = void 0;
zt.computeOperations = Gw;
var Mt;
(function(e) {
  e[e.COPY = 0] = "COPY", e[e.DOWNLOAD = 1] = "DOWNLOAD";
})(Mt || (zt.OperationKind = Mt = {}));
function Gw(e, t, n) {
  const r = cl(e.files), i = cl(t.files);
  let o = null;
  const a = t.files[0], s = [], c = a.name, p = r.get(c);
  if (p == null)
    throw new Error(`no file ${c} in old blockmap`);
  const l = i.get(c);
  let f = 0;
  const { checksumToOffset: h, checksumToOldSize: g } = Ww(r.get(c), p.offset, n);
  let _ = a.offset;
  for (let E = 0; E < l.checksums.length; _ += l.sizes[E], E++) {
    const T = l.sizes[E], A = l.checksums[E];
    let R = h.get(A);
    R != null && g.get(A) !== T && (n.warn(`Checksum ("${A}") matches, but size differs (old: ${g.get(A)}, new: ${T})`), R = void 0), R === void 0 ? (f++, o != null && o.kind === Mt.DOWNLOAD && o.end === _ ? o.end += T : (o = {
      kind: Mt.DOWNLOAD,
      start: _,
      end: _ + T
      // oldBlocks: null,
    }, ll(o, s, A, E))) : o != null && o.kind === Mt.COPY && o.end === R ? o.end += T : (o = {
      kind: Mt.COPY,
      start: R,
      end: R + T
      // oldBlocks: [checksum]
    }, ll(o, s, A, E));
  }
  return f > 0 && n.info(`File${a.name === "file" ? "" : " " + a.name} has ${f} changed blocks`), s;
}
const qw = process.env.DIFFERENTIAL_DOWNLOAD_PLAN_BUILDER_VALIDATE_RANGES === "true";
function ll(e, t, n, r) {
  if (qw && t.length !== 0) {
    const i = t[t.length - 1];
    if (i.kind === e.kind && e.start < i.end && e.start > i.start) {
      const o = [i.start, i.end, e.start, e.end].reduce((a, s) => a < s ? a : s);
      throw new Error(`operation (block index: ${r}, checksum: ${n}, kind: ${Mt[e.kind]}) overlaps previous operation (checksum: ${n}):
abs: ${i.start} until ${i.end} and ${e.start} until ${e.end}
rel: ${i.start - o} until ${i.end - o} and ${e.start - o} until ${e.end - o}`);
    }
  }
  t.push(e);
}
function Ww(e, t, n) {
  const r = /* @__PURE__ */ new Map(), i = /* @__PURE__ */ new Map();
  let o = t;
  for (let a = 0; a < e.checksums.length; a++) {
    const s = e.checksums[a], c = e.sizes[a], p = i.get(s);
    if (p === void 0)
      r.set(s, o), i.set(s, c);
    else if (n.debug != null) {
      const l = p === c ? "(same size)" : `(size: ${p}, this size: ${c})`;
      n.debug(`${s} duplicated in blockmap ${l}, it doesn't lead to broken differential downloader, just corresponding block will be skipped)`);
    }
    o += c;
  }
  return { checksumToOffset: r, checksumToOldSize: i };
}
function cl(e) {
  const t = /* @__PURE__ */ new Map();
  for (const n of e)
    t.set(n.name, n);
  return t;
}
Object.defineProperty(yn, "__esModule", { value: !0 });
yn.DataSplitter = void 0;
yn.copyData = nf;
const Fr = he, Vw = Ct, zw = Zn, Yw = zt, ul = Buffer.from(`\r
\r
`);
var gt;
(function(e) {
  e[e.INIT = 0] = "INIT", e[e.HEADER = 1] = "HEADER", e[e.BODY = 2] = "BODY";
})(gt || (gt = {}));
function nf(e, t, n, r, i) {
  const o = (0, Vw.createReadStream)("", {
    fd: n,
    autoClose: !1,
    start: e.start,
    // end is inclusive
    end: e.end - 1
  });
  o.on("error", r), o.once("end", i), o.pipe(t, {
    end: !1
  });
}
class Xw extends zw.Writable {
  constructor(t, n, r, i, o, a, s, c) {
    super(), this.out = t, this.options = n, this.partIndexToTaskIndex = r, this.partIndexToLength = o, this.finishHandler = a, this.grandTotalBytes = s, this.onProgress = c, this.start = Date.now(), this.nextUpdate = this.start + 1e3, this.transferred = 0, this.delta = 0, this.partIndex = -1, this.headerListBuffer = null, this.readState = gt.INIT, this.ignoreByteCount = 0, this.remainingPartDataCount = 0, this.actualPartLength = 0, this.boundaryLength = i.length + 4, this.ignoreByteCount = this.boundaryLength - 2;
  }
  get isFinished() {
    return this.partIndex === this.partIndexToLength.length;
  }
  // noinspection JSUnusedGlobalSymbols
  _write(t, n, r) {
    if (this.isFinished) {
      console.error(`Trailing ignored data: ${t.length} bytes`);
      return;
    }
    this.handleData(t).then(() => {
      if (this.onProgress) {
        const i = Date.now();
        (i >= this.nextUpdate || this.transferred === this.grandTotalBytes) && this.grandTotalBytes && (i - this.start) / 1e3 && (this.nextUpdate = i + 1e3, this.onProgress({
          total: this.grandTotalBytes,
          delta: this.delta,
          transferred: this.transferred,
          percent: this.transferred / this.grandTotalBytes * 100,
          bytesPerSecond: Math.round(this.transferred / ((i - this.start) / 1e3))
        }), this.delta = 0);
      }
      r();
    }).catch(r);
  }
  async handleData(t) {
    let n = 0;
    if (this.ignoreByteCount !== 0 && this.remainingPartDataCount !== 0)
      throw (0, Fr.newError)("Internal error", "ERR_DATA_SPLITTER_BYTE_COUNT_MISMATCH");
    if (this.ignoreByteCount > 0) {
      const r = Math.min(this.ignoreByteCount, t.length);
      this.ignoreByteCount -= r, n = r;
    } else if (this.remainingPartDataCount > 0) {
      const r = Math.min(this.remainingPartDataCount, t.length);
      this.remainingPartDataCount -= r, await this.processPartData(t, 0, r), n = r;
    }
    if (n !== t.length) {
      if (this.readState === gt.HEADER) {
        const r = this.searchHeaderListEnd(t, n);
        if (r === -1)
          return;
        n = r, this.readState = gt.BODY, this.headerListBuffer = null;
      }
      for (; ; ) {
        if (this.readState === gt.BODY)
          this.readState = gt.INIT;
        else {
          this.partIndex++;
          let a = this.partIndexToTaskIndex.get(this.partIndex);
          if (a == null)
            if (this.isFinished)
              a = this.options.end;
            else
              throw (0, Fr.newError)("taskIndex is null", "ERR_DATA_SPLITTER_TASK_INDEX_IS_NULL");
          const s = this.partIndex === 0 ? this.options.start : this.partIndexToTaskIndex.get(this.partIndex - 1) + 1;
          if (s < a)
            await this.copyExistingData(s, a);
          else if (s > a)
            throw (0, Fr.newError)("prevTaskIndex must be < taskIndex", "ERR_DATA_SPLITTER_TASK_INDEX_ASSERT_FAILED");
          if (this.isFinished) {
            this.onPartEnd(), this.finishHandler();
            return;
          }
          if (n = this.searchHeaderListEnd(t, n), n === -1) {
            this.readState = gt.HEADER;
            return;
          }
        }
        const r = this.partIndexToLength[this.partIndex], i = n + r, o = Math.min(i, t.length);
        if (await this.processPartStarted(t, n, o), this.remainingPartDataCount = r - (o - n), this.remainingPartDataCount > 0)
          return;
        if (n = i + this.boundaryLength, n >= t.length) {
          this.ignoreByteCount = this.boundaryLength - (t.length - i);
          return;
        }
      }
    }
  }
  copyExistingData(t, n) {
    return new Promise((r, i) => {
      const o = () => {
        if (t === n) {
          r();
          return;
        }
        const a = this.options.tasks[t];
        if (a.kind !== Yw.OperationKind.COPY) {
          i(new Error("Task kind must be COPY"));
          return;
        }
        nf(a, this.out, this.options.oldFileFd, i, () => {
          t++, o();
        });
      };
      o();
    });
  }
  searchHeaderListEnd(t, n) {
    const r = t.indexOf(ul, n);
    if (r !== -1)
      return r + ul.length;
    const i = n === 0 ? t : t.slice(n);
    return this.headerListBuffer == null ? this.headerListBuffer = i : this.headerListBuffer = Buffer.concat([this.headerListBuffer, i]), -1;
  }
  onPartEnd() {
    const t = this.partIndexToLength[this.partIndex - 1];
    if (this.actualPartLength !== t)
      throw (0, Fr.newError)(`Expected length: ${t} differs from actual: ${this.actualPartLength}`, "ERR_DATA_SPLITTER_LENGTH_MISMATCH");
    this.actualPartLength = 0;
  }
  processPartStarted(t, n, r) {
    return this.partIndex !== 0 && this.onPartEnd(), this.processPartData(t, n, r);
  }
  processPartData(t, n, r) {
    this.actualPartLength += r - n, this.transferred += r - n, this.delta += r - n;
    const i = this.out;
    return i.write(n === 0 && t.length === r ? t : t.slice(n, r)) ? Promise.resolve() : new Promise((o, a) => {
      i.on("error", a), i.once("drain", () => {
        i.removeListener("error", a), o();
      });
    });
  }
}
yn.DataSplitter = Xw;
var Ai = {};
Object.defineProperty(Ai, "__esModule", { value: !0 });
Ai.executeTasksUsingMultipleRangeRequests = Kw;
Ai.checkIsRangesSupported = Do;
const Io = he, fl = yn, dl = zt;
function Kw(e, t, n, r, i) {
  const o = (a) => {
    if (a >= t.length) {
      e.fileMetadataBuffer != null && n.write(e.fileMetadataBuffer), n.end();
      return;
    }
    const s = a + 1e3;
    Jw(e, {
      tasks: t,
      start: a,
      end: Math.min(t.length, s),
      oldFileFd: r
    }, n, () => o(s), i);
  };
  return o;
}
function Jw(e, t, n, r, i) {
  let o = "bytes=", a = 0, s = 0;
  const c = /* @__PURE__ */ new Map(), p = [];
  for (let h = t.start; h < t.end; h++) {
    const g = t.tasks[h];
    g.kind === dl.OperationKind.DOWNLOAD && (o += `${g.start}-${g.end - 1}, `, c.set(a, h), a++, p.push(g.end - g.start), s += g.end - g.start);
  }
  if (a <= 1) {
    const h = (g) => {
      if (g >= t.end) {
        r();
        return;
      }
      const _ = t.tasks[g++];
      if (_.kind === dl.OperationKind.COPY)
        (0, fl.copyData)(_, n, t.oldFileFd, i, () => h(g));
      else {
        const E = e.createRequestOptions();
        E.headers.Range = `bytes=${_.start}-${_.end - 1}`;
        const T = e.httpExecutor.createRequest(E, (A) => {
          A.on("error", i), Do(A, i) && (A.pipe(n, {
            end: !1
          }), A.once("end", () => h(g)));
        });
        e.httpExecutor.addErrorAndTimeoutHandlers(T, i), T.end();
      }
    };
    h(t.start);
    return;
  }
  const l = e.createRequestOptions();
  l.headers.Range = o.substring(0, o.length - 2);
  const f = e.httpExecutor.createRequest(l, (h) => {
    if (!Do(h, i))
      return;
    const g = (0, Io.safeGetHeader)(h, "content-type"), _ = /^multipart\/.+?\s*;\s*boundary=(?:"([^"]+)"|([^\s";]+))\s*$/i.exec(g);
    if (_ == null) {
      i(new Error(`Content-Type "multipart/byteranges" is expected, but got "${g}"`));
      return;
    }
    const E = new fl.DataSplitter(n, t, c, _[1] || _[2], p, r, s, e.options.onProgress);
    E.on("error", i), h.pipe(E), h.on("end", () => {
      setTimeout(() => {
        f.abort(), i(new Error("Response ends without calling any handlers"));
      }, 1e4);
    });
  });
  e.httpExecutor.addErrorAndTimeoutHandlers(f, i), f.end();
}
function Do(e, t) {
  if (e.statusCode >= 400)
    return t((0, Io.createHttpError)(e)), !1;
  if (e.statusCode !== 206) {
    const n = (0, Io.safeGetHeader)(e, "accept-ranges");
    if (n == null || n === "none")
      return t(new Error(`Server doesn't support Accept-Ranges (response code ${e.statusCode})`)), !1;
  }
  return !0;
}
var Si = {};
Object.defineProperty(Si, "__esModule", { value: !0 });
Si.ProgressDifferentialDownloadCallbackTransform = void 0;
const Qw = Zn;
var ln;
(function(e) {
  e[e.COPY = 0] = "COPY", e[e.DOWNLOAD = 1] = "DOWNLOAD";
})(ln || (ln = {}));
class Zw extends Qw.Transform {
  constructor(t, n, r) {
    super(), this.progressDifferentialDownloadInfo = t, this.cancellationToken = n, this.onProgress = r, this.start = Date.now(), this.transferred = 0, this.delta = 0, this.expectedBytes = 0, this.index = 0, this.operationType = ln.COPY, this.nextUpdate = this.start + 1e3;
  }
  _transform(t, n, r) {
    if (this.cancellationToken.cancelled) {
      r(new Error("cancelled"), null);
      return;
    }
    if (this.operationType == ln.COPY) {
      r(null, t);
      return;
    }
    this.transferred += t.length, this.delta += t.length;
    const i = Date.now();
    i >= this.nextUpdate && this.transferred !== this.expectedBytes && this.transferred !== this.progressDifferentialDownloadInfo.grandTotal && (this.nextUpdate = i + 1e3, this.onProgress({
      total: this.progressDifferentialDownloadInfo.grandTotal,
      delta: this.delta,
      transferred: this.transferred,
      percent: this.transferred / this.progressDifferentialDownloadInfo.grandTotal * 100,
      bytesPerSecond: Math.round(this.transferred / ((i - this.start) / 1e3))
    }), this.delta = 0), r(null, t);
  }
  beginFileCopy() {
    this.operationType = ln.COPY;
  }
  beginRangeDownload() {
    this.operationType = ln.DOWNLOAD, this.expectedBytes += this.progressDifferentialDownloadInfo.expectedByteCounts[this.index++];
  }
  endRangeDownload() {
    this.transferred !== this.progressDifferentialDownloadInfo.grandTotal && this.onProgress({
      total: this.progressDifferentialDownloadInfo.grandTotal,
      delta: this.delta,
      transferred: this.transferred,
      percent: this.transferred / this.progressDifferentialDownloadInfo.grandTotal * 100,
      bytesPerSecond: Math.round(this.transferred / ((Date.now() - this.start) / 1e3))
    });
  }
  // Called when we are 100% done with the connection/download
  _flush(t) {
    if (this.cancellationToken.cancelled) {
      t(new Error("cancelled"));
      return;
    }
    this.onProgress({
      total: this.progressDifferentialDownloadInfo.grandTotal,
      delta: this.delta,
      transferred: this.transferred,
      percent: 100,
      bytesPerSecond: Math.round(this.transferred / ((Date.now() - this.start) / 1e3))
    }), this.delta = 0, this.transferred = 0, t(null);
  }
}
Si.ProgressDifferentialDownloadCallbackTransform = Zw;
Object.defineProperty(ur, "__esModule", { value: !0 });
ur.DifferentialDownloader = void 0;
const $n = he, lo = Nt, ev = Ct, tv = yn, nv = bt, xr = zt, hl = Ai, rv = Si;
class iv {
  // noinspection TypeScriptAbstractClassConstructorCanBeMadeProtected
  constructor(t, n, r) {
    this.blockAwareFileInfo = t, this.httpExecutor = n, this.options = r, this.fileMetadataBuffer = null, this.logger = r.logger;
  }
  createRequestOptions() {
    const t = {
      headers: {
        ...this.options.requestHeaders,
        accept: "*/*"
      }
    };
    return (0, $n.configureRequestUrl)(this.options.newUrl, t), (0, $n.configureRequestOptions)(t), t;
  }
  doDownload(t, n) {
    if (t.version !== n.version)
      throw new Error(`version is different (${t.version} - ${n.version}), full download is required`);
    const r = this.logger, i = (0, xr.computeOperations)(t, n, r);
    r.debug != null && r.debug(JSON.stringify(i, null, 2));
    let o = 0, a = 0;
    for (const c of i) {
      const p = c.end - c.start;
      c.kind === xr.OperationKind.DOWNLOAD ? o += p : a += p;
    }
    const s = this.blockAwareFileInfo.size;
    if (o + a + (this.fileMetadataBuffer == null ? 0 : this.fileMetadataBuffer.length) !== s)
      throw new Error(`Internal error, size mismatch: downloadSize: ${o}, copySize: ${a}, newSize: ${s}`);
    return r.info(`Full: ${pl(s)}, To download: ${pl(o)} (${Math.round(o / (s / 100))}%)`), this.downloadFile(i);
  }
  downloadFile(t) {
    const n = [], r = () => Promise.all(n.map((i) => (0, lo.close)(i.descriptor).catch((o) => {
      this.logger.error(`cannot close file "${i.path}": ${o}`);
    })));
    return this.doDownloadFile(t, n).then(r).catch((i) => r().catch((o) => {
      try {
        this.logger.error(`cannot close files: ${o}`);
      } catch (a) {
        try {
          console.error(a);
        } catch {
        }
      }
      throw i;
    }).then(() => {
      throw i;
    }));
  }
  async doDownloadFile(t, n) {
    const r = await (0, lo.open)(this.options.oldFile, "r");
    n.push({ descriptor: r, path: this.options.oldFile });
    const i = await (0, lo.open)(this.options.newFile, "w");
    n.push({ descriptor: i, path: this.options.newFile });
    const o = (0, ev.createWriteStream)(this.options.newFile, { fd: i });
    await new Promise((a, s) => {
      const c = [];
      let p;
      if (!this.options.isUseMultipleRangeRequest && this.options.onProgress) {
        const A = [];
        let R = 0;
        for (const B of t)
          B.kind === xr.OperationKind.DOWNLOAD && (A.push(B.end - B.start), R += B.end - B.start);
        const P = {
          expectedByteCounts: A,
          grandTotal: R
        };
        p = new rv.ProgressDifferentialDownloadCallbackTransform(P, this.options.cancellationToken, this.options.onProgress), c.push(p);
      }
      const l = new $n.DigestTransform(this.blockAwareFileInfo.sha512);
      l.isValidateOnEnd = !1, c.push(l), o.on("finish", () => {
        o.close(() => {
          n.splice(1, 1);
          try {
            l.validate();
          } catch (A) {
            s(A);
            return;
          }
          a(void 0);
        });
      }), c.push(o);
      let f = null;
      for (const A of c)
        A.on("error", s), f == null ? f = A : f = f.pipe(A);
      const h = c[0];
      let g;
      if (this.options.isUseMultipleRangeRequest) {
        g = (0, hl.executeTasksUsingMultipleRangeRequests)(this, t, h, r, s), g(0);
        return;
      }
      let _ = 0, E = null;
      this.logger.info(`Differential download: ${this.options.newUrl}`);
      const T = this.createRequestOptions();
      T.redirect = "manual", g = (A) => {
        var R, P;
        if (A >= t.length) {
          this.fileMetadataBuffer != null && h.write(this.fileMetadataBuffer), h.end();
          return;
        }
        const B = t[A++];
        if (B.kind === xr.OperationKind.COPY) {
          p && p.beginFileCopy(), (0, tv.copyData)(B, h, r, s, () => g(A));
          return;
        }
        const q = `bytes=${B.start}-${B.end - 1}`;
        T.headers.range = q, (P = (R = this.logger) === null || R === void 0 ? void 0 : R.debug) === null || P === void 0 || P.call(R, `download range: ${q}`), p && p.beginRangeDownload();
        const Q = this.httpExecutor.createRequest(T, (Z) => {
          Z.on("error", s), Z.on("aborted", () => {
            s(new Error("response has been aborted by the server"));
          }), Z.statusCode >= 400 && s((0, $n.createHttpError)(Z)), Z.pipe(h, {
            end: !1
          }), Z.once("end", () => {
            p && p.endRangeDownload(), ++_ === 100 ? (_ = 0, setTimeout(() => g(A), 1e3)) : g(A);
          });
        });
        Q.on("redirect", (Z, ae, U) => {
          this.logger.info(`Redirect to ${ov(U)}`), E = U, (0, $n.configureRequestUrl)(new nv.URL(E), T), Q.followRedirect();
        }), this.httpExecutor.addErrorAndTimeoutHandlers(Q, s), Q.end();
      }, g(0);
    });
  }
  async readRemoteBytes(t, n) {
    const r = Buffer.allocUnsafe(n + 1 - t), i = this.createRequestOptions();
    i.headers.range = `bytes=${t}-${n}`;
    let o = 0;
    if (await this.request(i, (a) => {
      a.copy(r, o), o += a.length;
    }), o !== r.length)
      throw new Error(`Received data length ${o} is not equal to expected ${r.length}`);
    return r;
  }
  request(t, n) {
    return new Promise((r, i) => {
      const o = this.httpExecutor.createRequest(t, (a) => {
        (0, hl.checkIsRangesSupported)(a, i) && (a.on("error", i), a.on("aborted", () => {
          i(new Error("response has been aborted by the server"));
        }), a.on("data", n), a.on("end", () => r()));
      });
      this.httpExecutor.addErrorAndTimeoutHandlers(o, i), o.end();
    });
  }
}
ur.DifferentialDownloader = iv;
function pl(e, t = " KB") {
  return new Intl.NumberFormat("en").format((e / 1024).toFixed(2)) + t;
}
function ov(e) {
  const t = e.indexOf("?");
  return t < 0 ? e : e.substring(0, t);
}
Object.defineProperty(Ti, "__esModule", { value: !0 });
Ti.GenericDifferentialDownloader = void 0;
const av = ur;
class sv extends av.DifferentialDownloader {
  download(t, n) {
    return this.doDownload(t, n);
  }
}
Ti.GenericDifferentialDownloader = sv;
var Rt = {};
(function(e) {
  Object.defineProperty(e, "__esModule", { value: !0 }), e.UpdaterSignal = e.UPDATE_DOWNLOADED = e.DOWNLOAD_PROGRESS = e.CancellationToken = void 0, e.addHandler = r;
  const t = he;
  Object.defineProperty(e, "CancellationToken", { enumerable: !0, get: function() {
    return t.CancellationToken;
  } }), e.DOWNLOAD_PROGRESS = "download-progress", e.UPDATE_DOWNLOADED = "update-downloaded";
  class n {
    constructor(o) {
      this.emitter = o;
    }
    /**
     * Emitted when an authenticating proxy is [asking for user credentials](https://github.com/electron/electron/blob/master/docs/api/client-request.md#event-login).
     */
    login(o) {
      r(this.emitter, "login", o);
    }
    progress(o) {
      r(this.emitter, e.DOWNLOAD_PROGRESS, o);
    }
    updateDownloaded(o) {
      r(this.emitter, e.UPDATE_DOWNLOADED, o);
    }
    updateCancelled(o) {
      r(this.emitter, "update-cancelled", o);
    }
  }
  e.UpdaterSignal = n;
  function r(i, o, a) {
    i.on(o, a);
  }
})(Rt);
Object.defineProperty(vt, "__esModule", { value: !0 });
vt.NoOpLogger = vt.AppUpdater = void 0;
const Re = he, lv = er, cv = ti, uv = kl, Ve = Nt, fv = we, co = ui, He = ee, xt = zu, ml = lr, dv = gi, gl = Yu, hv = cr, uo = Ei, fo = Bl, pv = Ti, Qt = Rt;
class ua extends uv.EventEmitter {
  /**
   * Get the update channel. Doesn't return `channel` from the update configuration, only if was previously set.
   */
  get channel() {
    return this._channel;
  }
  /**
   * Set the update channel. Overrides `channel` in the update configuration.
   *
   * `allowDowngrade` will be automatically set to `true`. If this behavior is not suitable for you, simple set `allowDowngrade` explicitly after.
   */
  set channel(t) {
    if (this._channel != null) {
      if (typeof t != "string")
        throw (0, Re.newError)(`Channel must be a string, but got: ${t}`, "ERR_UPDATER_INVALID_CHANNEL");
      if (t.length === 0)
        throw (0, Re.newError)("Channel must be not an empty string", "ERR_UPDATER_INVALID_CHANNEL");
    }
    this._channel = t, this.allowDowngrade = !0;
  }
  /**
   *  Shortcut for explicitly adding auth tokens to request headers
   */
  addAuthHeader(t) {
    this.requestHeaders = Object.assign({}, this.requestHeaders, {
      authorization: t
    });
  }
  // noinspection JSMethodCanBeStatic,JSUnusedGlobalSymbols
  get netSession() {
    return (0, gl.getNetSession)();
  }
  /**
   * The logger. You can pass [electron-log](https://github.com/megahertz/electron-log), [winston](https://github.com/winstonjs/winston) or another logger with the following interface: `{ info(), warn(), error() }`.
   * Set it to `null` if you would like to disable a logging feature.
   */
  get logger() {
    return this._logger;
  }
  set logger(t) {
    this._logger = t ?? new rf();
  }
  // noinspection JSUnusedGlobalSymbols
  /**
   * test only
   * @private
   */
  set updateConfigPath(t) {
    this.clientPromise = null, this._appUpdateConfigPath = t, this.configOnDisk = new co.Lazy(() => this.loadUpdateConfig());
  }
  /**
   * Allows developer to override default logic for determining if an update is supported.
   * The default logic compares the `UpdateInfo` minimum system version against the `os.release()` with `semver` package
   */
  get isUpdateSupported() {
    return this._isUpdateSupported;
  }
  set isUpdateSupported(t) {
    t && (this._isUpdateSupported = t);
  }
  /**
   * Allows developer to override default logic for determining if the user is below the rollout threshold.
   * The default logic compares the staging percentage with numerical representation of user ID.
   * An override can define custom logic, or bypass it if needed.
   */
  get isUserWithinRollout() {
    return this._isUserWithinRollout;
  }
  set isUserWithinRollout(t) {
    t && (this._isUserWithinRollout = t);
  }
  constructor(t, n) {
    super(), this.autoDownload = !0, this.autoInstallOnAppQuit = !0, this.autoRunAppAfterInstall = !0, this.allowPrerelease = !1, this.fullChangelog = !1, this.allowDowngrade = !1, this.disableWebInstaller = !1, this.disableDifferentialDownload = !1, this.forceDevUpdateConfig = !1, this.previousBlockmapBaseUrlOverride = null, this._channel = null, this.downloadedUpdateHelper = null, this.requestHeaders = null, this._logger = console, this.signals = new Qt.UpdaterSignal(this), this._appUpdateConfigPath = null, this._isUpdateSupported = (o) => this.checkIfUpdateSupported(o), this._isUserWithinRollout = (o) => this.isStagingMatch(o), this.clientPromise = null, this.stagingUserIdPromise = new co.Lazy(() => this.getOrCreateStagingUserId()), this.configOnDisk = new co.Lazy(() => this.loadUpdateConfig()), this.checkForUpdatesPromise = null, this.downloadPromise = null, this.updateInfoAndProvider = null, this._testOnlyOptions = null, this.on("error", (o) => {
      this._logger.error(`Error: ${o.stack || o.message}`);
    }), n == null ? (this.app = new dv.ElectronAppAdapter(), this.httpExecutor = new gl.ElectronHttpExecutor((o, a) => this.emit("login", o, a))) : (this.app = n, this.httpExecutor = null);
    const r = this.app.version, i = (0, xt.parse)(r);
    if (i == null)
      throw (0, Re.newError)(`App version is not a valid semver version: "${r}"`, "ERR_UPDATER_INVALID_VERSION");
    this.currentVersion = i, this.allowPrerelease = mv(i), t != null && (this.setFeedURL(t), typeof t != "string" && t.requestHeaders && (this.requestHeaders = t.requestHeaders));
  }
  //noinspection JSMethodCanBeStatic,JSUnusedGlobalSymbols
  getFeedURL() {
    return "Deprecated. Do not use it.";
  }
  /**
   * Configure update provider. If value is `string`, [GenericServerOptions](https://www.electron.build/publish#genericserveroptions) will be set with value as `url`.
   * @param options If you want to override configuration in the `app-update.yml`.
   */
  setFeedURL(t) {
    const n = this.createProviderRuntimeOptions();
    let r;
    typeof t == "string" ? r = new hv.GenericProvider({ provider: "generic", url: t }, this, {
      ...n,
      isUseMultipleRangeRequest: (0, uo.isUrlProbablySupportMultiRangeRequests)(t)
    }) : r = (0, uo.createClient)(t, this, n), this.clientPromise = Promise.resolve(r);
  }
  /**
   * Asks the server whether there is an update.
   * @returns null if the updater is disabled, otherwise info about the latest version
   */
  checkForUpdates() {
    if (!this.isUpdaterActive())
      return Promise.resolve(null);
    let t = this.checkForUpdatesPromise;
    if (t != null)
      return this._logger.info("Checking for update (already in progress)"), t;
    const n = () => this.checkForUpdatesPromise = null;
    return this._logger.info("Checking for update"), t = this.doCheckForUpdates().then((r) => (n(), r)).catch((r) => {
      throw n(), this.emit("error", r, `Cannot check for updates: ${(r.stack || r).toString()}`), r;
    }), this.checkForUpdatesPromise = t, t;
  }
  isUpdaterActive() {
    return this.app.isPackaged || this.forceDevUpdateConfig ? !0 : (this._logger.info("Skip checkForUpdates because application is not packed and dev update config is not forced"), !1);
  }
  // noinspection JSUnusedGlobalSymbols
  checkForUpdatesAndNotify(t) {
    return this.checkForUpdates().then((n) => n != null && n.downloadPromise ? (n.downloadPromise.then(() => {
      const r = ua.formatDownloadNotification(n.updateInfo.version, this.app.name, t);
      new Ht.Notification(r).show();
    }), n) : (this._logger.debug != null && this._logger.debug("checkForUpdatesAndNotify called, downloadPromise is null"), n));
  }
  static formatDownloadNotification(t, n, r) {
    return r == null && (r = {
      title: "A new update is ready to install",
      body: "{appName} version {version} has been downloaded and will be automatically installed on exit"
    }), r = {
      title: r.title.replace("{appName}", n).replace("{version}", t),
      body: r.body.replace("{appName}", n).replace("{version}", t)
    }, r;
  }
  async isStagingMatch(t) {
    const n = t.stagingPercentage;
    let r = n;
    if (r == null)
      return !0;
    if (r = parseInt(r, 10), isNaN(r))
      return this._logger.warn(`Staging percentage is NaN: ${n}`), !0;
    r = r / 100;
    const i = await this.stagingUserIdPromise.value, a = Re.UUID.parse(i).readUInt32BE(12) / 4294967295;
    return this._logger.info(`Staging percentage: ${r}, percentage: ${a}, user id: ${i}`), a < r;
  }
  computeFinalHeaders(t) {
    return this.requestHeaders != null && Object.assign(t, this.requestHeaders), t;
  }
  async isUpdateAvailable(t) {
    const n = (0, xt.parse)(t.version);
    if (n == null)
      throw (0, Re.newError)(`This file could not be downloaded, or the latest version (from update server) does not have a valid semver version: "${t.version}"`, "ERR_UPDATER_INVALID_VERSION");
    const r = this.currentVersion;
    if ((0, xt.eq)(n, r) || !await Promise.resolve(this.isUpdateSupported(t)) || !await Promise.resolve(this.isUserWithinRollout(t)))
      return !1;
    const o = (0, xt.gt)(n, r), a = (0, xt.lt)(n, r);
    return o ? !0 : this.allowDowngrade && a;
  }
  checkIfUpdateSupported(t) {
    const n = t == null ? void 0 : t.minimumSystemVersion, r = (0, cv.release)();
    if (n)
      try {
        if ((0, xt.lt)(r, n))
          return this._logger.info(`Current OS version ${r} is less than the minimum OS version required ${n} for version ${r}`), !1;
      } catch (i) {
        this._logger.warn(`Failed to compare current OS version(${r}) with minimum OS version(${n}): ${(i.message || i).toString()}`);
      }
    return !0;
  }
  async getUpdateInfoAndProvider() {
    await this.app.whenReady(), this.clientPromise == null && (this.clientPromise = this.configOnDisk.value.then((r) => (0, uo.createClient)(r, this, this.createProviderRuntimeOptions())));
    const t = await this.clientPromise, n = await this.stagingUserIdPromise.value;
    return t.setRequestHeaders(this.computeFinalHeaders({ "x-user-staging-id": n })), {
      info: await t.getLatestVersion(),
      provider: t
    };
  }
  createProviderRuntimeOptions() {
    return {
      isUseMultipleRangeRequest: !0,
      platform: this._testOnlyOptions == null ? process.platform : this._testOnlyOptions.platform,
      executor: this.httpExecutor
    };
  }
  async doCheckForUpdates() {
    this.emit("checking-for-update");
    const t = await this.getUpdateInfoAndProvider(), n = t.info;
    if (!await this.isUpdateAvailable(n))
      return this._logger.info(`Update for version ${this.currentVersion.format()} is not available (latest version: ${n.version}, downgrade is ${this.allowDowngrade ? "allowed" : "disallowed"}).`), this.emit("update-not-available", n), {
        isUpdateAvailable: !1,
        versionInfo: n,
        updateInfo: n
      };
    this.updateInfoAndProvider = t, this.onUpdateAvailable(n);
    const r = new Re.CancellationToken();
    return {
      isUpdateAvailable: !0,
      versionInfo: n,
      updateInfo: n,
      cancellationToken: r,
      downloadPromise: this.autoDownload ? this.downloadUpdate(r) : null
    };
  }
  onUpdateAvailable(t) {
    this._logger.info(`Found version ${t.version} (url: ${(0, Re.asArray)(t.files).map((n) => n.url).join(", ")})`), this.emit("update-available", t);
  }
  /**
   * Start downloading update manually. You can use this method if `autoDownload` option is set to `false`.
   * @returns {Promise<Array<string>>} Paths to downloaded files.
   */
  downloadUpdate(t = new Re.CancellationToken()) {
    const n = this.updateInfoAndProvider;
    if (n == null) {
      const i = new Error("Please check update first");
      return this.dispatchError(i), Promise.reject(i);
    }
    if (this.downloadPromise != null)
      return this._logger.info("Downloading update (already in progress)"), this.downloadPromise;
    this._logger.info(`Downloading update from ${(0, Re.asArray)(n.info.files).map((i) => i.url).join(", ")}`);
    const r = (i) => {
      if (!(i instanceof Re.CancellationError))
        try {
          this.dispatchError(i);
        } catch (o) {
          this._logger.warn(`Cannot dispatch error event: ${o.stack || o}`);
        }
      return i;
    };
    return this.downloadPromise = this.doDownloadUpdate({
      updateInfoAndProvider: n,
      requestHeaders: this.computeRequestHeaders(n.provider),
      cancellationToken: t,
      disableWebInstaller: this.disableWebInstaller,
      disableDifferentialDownload: this.disableDifferentialDownload
    }).catch((i) => {
      throw r(i);
    }).finally(() => {
      this.downloadPromise = null;
    }), this.downloadPromise;
  }
  dispatchError(t) {
    this.emit("error", t, (t.stack || t).toString());
  }
  dispatchUpdateDownloaded(t) {
    this.emit(Qt.UPDATE_DOWNLOADED, t);
  }
  async loadUpdateConfig() {
    return this._appUpdateConfigPath == null && (this._appUpdateConfigPath = this.app.appUpdateConfigPath), (0, fv.load)(await (0, Ve.readFile)(this._appUpdateConfigPath, "utf-8"));
  }
  computeRequestHeaders(t) {
    const n = t.fileExtraDownloadHeaders;
    if (n != null) {
      const r = this.requestHeaders;
      return r == null ? n : {
        ...n,
        ...r
      };
    }
    return this.computeFinalHeaders({ accept: "*/*" });
  }
  async getOrCreateStagingUserId() {
    const t = He.join(this.app.userDataPath, ".updaterId");
    try {
      const r = await (0, Ve.readFile)(t, "utf-8");
      if (Re.UUID.check(r))
        return r;
      this._logger.warn(`Staging user id file exists, but content was invalid: ${r}`);
    } catch (r) {
      r.code !== "ENOENT" && this._logger.warn(`Couldn't read staging user ID, creating a blank one: ${r}`);
    }
    const n = Re.UUID.v5((0, lv.randomBytes)(4096), Re.UUID.OID);
    this._logger.info(`Generated new staging user ID: ${n}`);
    try {
      await (0, Ve.outputFile)(t, n);
    } catch (r) {
      this._logger.warn(`Couldn't write out staging user ID: ${r}`);
    }
    return n;
  }
  /** @internal */
  get isAddNoCacheQuery() {
    const t = this.requestHeaders;
    if (t == null)
      return !0;
    for (const n of Object.keys(t)) {
      const r = n.toLowerCase();
      if (r === "authorization" || r === "private-token")
        return !1;
    }
    return !0;
  }
  async getOrCreateDownloadHelper() {
    let t = this.downloadedUpdateHelper;
    if (t == null) {
      const n = (await this.configOnDisk.value).updaterCacheDirName, r = this._logger;
      n == null && r.error("updaterCacheDirName is not specified in app-update.yml Was app build using at least electron-builder 20.34.0?");
      const i = He.join(this.app.baseCachePath, n || this.app.name);
      r.debug != null && r.debug(`updater cache dir: ${i}`), t = new ml.DownloadedUpdateHelper(i), this.downloadedUpdateHelper = t;
    }
    return t;
  }
  async executeDownload(t) {
    const n = t.fileInfo, r = {
      headers: t.downloadUpdateOptions.requestHeaders,
      cancellationToken: t.downloadUpdateOptions.cancellationToken,
      sha2: n.info.sha2,
      sha512: n.info.sha512
    };
    this.listenerCount(Qt.DOWNLOAD_PROGRESS) > 0 && (r.onProgress = (R) => this.emit(Qt.DOWNLOAD_PROGRESS, R));
    const i = t.downloadUpdateOptions.updateInfoAndProvider.info, o = i.version, a = n.packageInfo;
    function s() {
      const R = decodeURIComponent(t.fileInfo.url.pathname);
      return R.toLowerCase().endsWith(`.${t.fileExtension.toLowerCase()}`) ? He.basename(R) : He.basename(t.fileInfo.info.url);
    }
    const c = await this.getOrCreateDownloadHelper(), p = c.cacheDirForPendingUpdate;
    await (0, Ve.mkdir)(p, { recursive: !0 });
    const l = s();
    let f = He.join(p, l);
    const h = a == null ? null : He.join(p, `package-${o}${He.extname(a.path) || ".7z"}`), g = async (R) => {
      await c.setDownloadedFile(f, h, i, n, l, R), await t.done({
        ...i,
        downloadedFile: f
      });
      const P = He.join(p, "current.blockmap");
      return await (0, Ve.pathExists)(P) && await (0, Ve.copyFile)(P, He.join(c.cacheDir, "current.blockmap")), h == null ? [f] : [f, h];
    }, _ = this._logger, E = await c.validateDownloadedPath(f, i, n, _);
    if (E != null)
      return f = E, await g(!1);
    const T = async () => (await c.clear().catch(() => {
    }), await (0, Ve.unlink)(f).catch(() => {
    })), A = await (0, ml.createTempUpdateFile)(`temp-${l}`, p, _);
    try {
      await t.task(A, r, h, T), await (0, Re.retry)(() => (0, Ve.rename)(A, f), {
        retries: 60,
        interval: 500,
        shouldRetry: (R) => R instanceof Error && /^EBUSY:/.test(R.message) ? !0 : (_.warn(`Cannot rename temp file to final file: ${R.message || R.stack}`), !1)
      });
    } catch (R) {
      throw await T(), R instanceof Re.CancellationError && (_.info("cancelled"), this.emit("update-cancelled", i)), R;
    }
    return _.info(`New version ${o} has been downloaded to ${f}`), await g(!0);
  }
  async differentialDownloadInstaller(t, n, r, i, o) {
    try {
      if (this._testOnlyOptions != null && !this._testOnlyOptions.isUseDifferentialDownload)
        return !0;
      const a = n.updateInfoAndProvider.provider, s = await a.getBlockMapFiles(t.url, this.app.version, n.updateInfoAndProvider.info.version, this.previousBlockmapBaseUrlOverride);
      this._logger.info(`Download block maps (old: "${s[0]}", new: ${s[1]})`);
      const c = async (_) => {
        const E = await this.httpExecutor.downloadToBuffer(_, {
          headers: n.requestHeaders,
          cancellationToken: n.cancellationToken
        });
        if (E == null || E.length === 0)
          throw new Error(`Blockmap "${_.href}" is empty`);
        try {
          return JSON.parse((0, fo.gunzipSync)(E).toString());
        } catch (T) {
          throw new Error(`Cannot parse blockmap "${_.href}", error: ${T}`);
        }
      }, p = {
        newUrl: t.url,
        oldFile: He.join(this.downloadedUpdateHelper.cacheDir, o),
        logger: this._logger,
        newFile: r,
        isUseMultipleRangeRequest: a.isUseMultipleRangeRequest,
        requestHeaders: n.requestHeaders,
        cancellationToken: n.cancellationToken
      };
      this.listenerCount(Qt.DOWNLOAD_PROGRESS) > 0 && (p.onProgress = (_) => this.emit(Qt.DOWNLOAD_PROGRESS, _));
      const l = async (_, E) => {
        const T = He.join(E, "current.blockmap");
        await (0, Ve.outputFile)(T, (0, fo.gzipSync)(JSON.stringify(_)));
      }, f = async (_) => {
        const E = He.join(_, "current.blockmap");
        try {
          if (await (0, Ve.pathExists)(E))
            return JSON.parse((0, fo.gunzipSync)(await (0, Ve.readFile)(E)).toString());
        } catch (T) {
          this._logger.warn(`Cannot parse blockmap "${E}", error: ${T}`);
        }
        return null;
      }, h = await c(s[1]);
      await l(h, this.downloadedUpdateHelper.cacheDirForPendingUpdate);
      let g = await f(this.downloadedUpdateHelper.cacheDir);
      return g == null && (g = await c(s[0])), await new pv.GenericDifferentialDownloader(t.info, this.httpExecutor, p).download(g, h), !1;
    } catch (a) {
      if (this._logger.error(`Cannot download differentially, fallback to full download: ${a.stack || a}`), this._testOnlyOptions != null)
        throw a;
      return !0;
    }
  }
}
vt.AppUpdater = ua;
function mv(e) {
  const t = (0, xt.prerelease)(e);
  return t != null && t.length > 0;
}
class rf {
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  info(t) {
  }
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  warn(t) {
  }
  // eslint-disable-next-line @typescript-eslint/no-unused-vars
  error(t) {
  }
}
vt.NoOpLogger = rf;
Object.defineProperty(Wt, "__esModule", { value: !0 });
Wt.BaseUpdater = void 0;
const El = ei, ho = ee, gv = vt;
class Ev extends gv.AppUpdater {
  constructor(t, n) {
    super(t, n), this.quitAndInstallCalled = !1, this.quitHandlerAdded = !1;
  }
  quitAndInstall(t = !1, n = !1) {
    this._logger.info("Install on explicit quitAndInstall"), this.install(t, t ? n : this.autoRunAppAfterInstall) ? setImmediate(() => {
      Ht.autoUpdater.emit("before-quit-for-update"), this.app.quit();
    }) : this.quitAndInstallCalled = !1;
  }
  executeDownload(t) {
    return super.executeDownload({
      ...t,
      done: (n) => (this.dispatchUpdateDownloaded(n), this.addQuitHandler(), Promise.resolve())
    });
  }
  get installerPath() {
    return this.downloadedUpdateHelper == null ? null : this.downloadedUpdateHelper.file;
  }
  // must be sync (because quit even handler is not async)
  install(t = !1, n = !1) {
    if (this.quitAndInstallCalled)
      return this._logger.warn("install call ignored: quitAndInstallCalled is set to true"), !1;
    const r = this.downloadedUpdateHelper, i = this.installerPath, o = r == null ? null : r.downloadedFileInfo;
    if (i == null || o == null)
      return this.dispatchError(new Error("No update filepath provided, can't quit and install")), !1;
    this.quitAndInstallCalled = !0;
    try {
      return this._logger.info(`Install: isSilent: ${t}, isForceRunAfter: ${n}`), this.doInstall({
        isSilent: t,
        isForceRunAfter: n,
        isAdminRightsRequired: o.isAdminRightsRequired
      });
    } catch (a) {
      return this.dispatchError(a), !1;
    }
  }
  addQuitHandler() {
    this.quitHandlerAdded || !this.autoInstallOnAppQuit || (this.quitHandlerAdded = !0, this.app.onQuit((t) => {
      if (this.quitAndInstallCalled) {
        this._logger.info("Update installer has already been triggered. Quitting application.");
        return;
      }
      if (!this.autoInstallOnAppQuit) {
        this._logger.info("Update will not be installed on quit because autoInstallOnAppQuit is set to false.");
        return;
      }
      if (t !== 0) {
        this._logger.info(`Update will be not installed on quit because application is quitting with exit code ${t}`);
        return;
      }
      this._logger.info("Auto install update on quit"), this.install(!0, !1);
    }));
  }
  /**
   * Strips relative-path entries from a PATH string.
   * Prevents PATH-poisoning where a writable directory earlier in PATH shadows
   * a trusted package manager binary.
   */
  sanitizeEnvPath(t) {
    return t.split(ho.delimiter).filter((n) => ho.isAbsolute(n)).join(ho.delimiter);
  }
  spawnSyncLog(t, n = [], r = {}) {
    var i;
    this._logger.info(`Executing: ${t} with args: ${n}`);
    const o = { ...process.env, ...r }, a = (0, El.spawnSync)(t, n, {
      env: { ...o, PATH: this.sanitizeEnvPath((i = o.PATH) !== null && i !== void 0 ? i : "") },
      encoding: "utf-8",
      shell: !0
    }), { error: s, status: c, stdout: p, stderr: l } = a;
    if (s != null)
      throw this._logger.error(l), s;
    if (c != null && c !== 0)
      throw this._logger.error(l), new Error(`Command ${t} exited with code ${c}`);
    return p.trim();
  }
  /**
   * This handles both node 8 and node 10 way of emitting error when spawning a process
   *   - node 8: Throws the error
   *   - node 10: Emit the error(Need to listen with on)
   */
  // https://github.com/electron-userland/electron-builder/issues/1129
  // Node 8 sends errors: https://nodejs.org/dist/latest-v8.x/docs/api/errors.html#errors_common_system_errors
  async spawnLog(t, n = [], r = void 0, i = "ignore") {
    return this._logger.info(`Executing: ${t} with args: ${n}`), new Promise((o, a) => {
      try {
        const s = { stdio: i, env: r, detached: !0 }, c = (0, El.spawn)(t, n, s);
        c.on("error", (p) => {
          a(p);
        }), c.unref(), c.pid !== void 0 && o(!0);
      } catch (s) {
        a(s);
      }
    });
  }
}
Wt.BaseUpdater = Ev;
var zn = {}, fr = {};
Object.defineProperty(fr, "__esModule", { value: !0 });
fr.FileWithEmbeddedBlockMapDifferentialDownloader = void 0;
const Zt = Nt, yv = ur, _v = Bl;
class wv extends yv.DifferentialDownloader {
  async download() {
    const t = this.blockAwareFileInfo, n = t.size, r = n - (t.blockMapSize + 4);
    this.fileMetadataBuffer = await this.readRemoteBytes(r, n - 1);
    const i = of(this.fileMetadataBuffer.slice(0, this.fileMetadataBuffer.length - 4));
    await this.doDownload(await vv(this.options.oldFile), i);
  }
}
fr.FileWithEmbeddedBlockMapDifferentialDownloader = wv;
function of(e) {
  return JSON.parse((0, _v.inflateRawSync)(e).toString());
}
async function vv(e) {
  const t = await (0, Zt.open)(e, "r");
  try {
    const n = (await (0, Zt.fstat)(t)).size, r = Buffer.allocUnsafe(4);
    await (0, Zt.read)(t, r, 0, r.length, n - r.length);
    const i = Buffer.allocUnsafe(r.readUInt32BE(0));
    return await (0, Zt.read)(t, i, 0, i.length, n - r.length - i.length), await (0, Zt.close)(t), of(i);
  } catch (n) {
    throw await (0, Zt.close)(t), n;
  }
}
Object.defineProperty(zn, "__esModule", { value: !0 });
zn.AppImageUpdater = void 0;
const po = he, yl = ei, Tv = Nt, Av = Ct, en = ee, Sv = Wt, Cv = fr, bv = ce, _l = Rt;
class Nv extends Sv.BaseUpdater {
  constructor(t, n) {
    super(t, n);
  }
  isUpdaterActive() {
    return process.env.APPIMAGE == null && !this.forceDevUpdateConfig ? (process.env.SNAP == null ? this._logger.warn("APPIMAGE env is not defined, current application is not an AppImage") : this._logger.info("SNAP env is defined, updater is disabled"), !1) : super.isUpdaterActive();
  }
  /*** @private */
  doDownloadUpdate(t) {
    const n = t.updateInfoAndProvider.provider, r = (0, bv.findFile)(n.resolveFiles(t.updateInfoAndProvider.info), "AppImage", ["rpm", "deb", "pacman"]);
    return this.executeDownload({
      fileExtension: "AppImage",
      fileInfo: r,
      downloadUpdateOptions: t,
      task: async (i, o) => {
        const a = process.env.APPIMAGE;
        if (a == null)
          throw (0, po.newError)("APPIMAGE env is not defined", "ERR_UPDATER_OLD_FILE_NOT_FOUND");
        (t.disableDifferentialDownload || await this.downloadDifferential(r, a, i, n, t)) && await this.httpExecutor.download(r.url, i, o), await (0, Tv.chmod)(i, 493);
      }
    });
  }
  async downloadDifferential(t, n, r, i, o) {
    try {
      const a = {
        newUrl: t.url,
        oldFile: n,
        logger: this._logger,
        newFile: r,
        isUseMultipleRangeRequest: i.isUseMultipleRangeRequest,
        requestHeaders: o.requestHeaders,
        cancellationToken: o.cancellationToken
      };
      return this.listenerCount(_l.DOWNLOAD_PROGRESS) > 0 && (a.onProgress = (s) => this.emit(_l.DOWNLOAD_PROGRESS, s)), await new Cv.FileWithEmbeddedBlockMapDifferentialDownloader(t.info, this.httpExecutor, a).download(), !1;
    } catch (a) {
      return this._logger.error(`Cannot download differentially, fallback to full download: ${a.stack || a}`), process.platform === "linux";
    }
  }
  doInstall(t) {
    const n = process.env.APPIMAGE;
    if (n == null)
      throw (0, po.newError)("APPIMAGE env is not defined", "ERR_UPDATER_OLD_FILE_NOT_FOUND");
    if (!en.isAbsolute(n) || n.includes("\0"))
      throw (0, po.newError)(`APPIMAGE env is not a valid absolute path: "${n}"`, "ERR_UPDATER_OLD_FILE_NOT_FOUND");
    (0, Av.unlinkSync)(n);
    let r;
    const i = en.basename(n), o = this.installerPath;
    if (o == null)
      return this.dispatchError(new Error("No update filepath provided, can't quit and install")), !1;
    en.basename(o) === i || !/\d+\.\d+\.\d+/.test(i) ? r = n : r = en.join(en.dirname(n), en.basename(o)), (0, yl.execFileSync)("mv", ["-f", o, r]), r !== n && this.emit("appimage-filename-updated", r);
    const a = {
      ...process.env,
      APPIMAGE_SILENT_INSTALL: "true"
    };
    return t.isForceRunAfter ? this.spawnLog(r, [], a) : (a.APPIMAGE_EXIT_AFTER_INSTALL = "true", (0, yl.execFileSync)(r, [], { env: a })), !0;
  }
}
zn.AppImageUpdater = Nv;
var Yn = {}, _n = {};
Object.defineProperty(_n, "__esModule", { value: !0 });
_n.LinuxUpdater = void 0;
const Rv = Wt, $v = /^[a-zA-Z0-9_-]+$/;
class Ov extends Rv.BaseUpdater {
  constructor(t, n) {
    super(t, n);
  }
  /**
   * Returns true if the current process is running as root.
   */
  isRunningAsRoot() {
    var t;
    return ((t = process.getuid) === null || t === void 0 ? void 0 : t.call(process)) === 0;
  }
  /**
   * Sanitizes the installer path for use with shell:true spawn calls.
   * Backslash-escapes metacharacters that have special meaning in POSIX shell.
   * Note: paths containing single-quotes (') are not supported.
   */
  get installerPath() {
    const t = super.installerPath;
    return t == null ? null : t.replace(/\\/g, "\\\\").replace(/([`$!" ;|&()<>])/g, "\\$1").replace(/[\n\r]/g, "");
  }
  runCommandWithSudoIfNeeded(t) {
    if (this.isRunningAsRoot())
      return this._logger.info("Running as root, no need to use sudo"), this.spawnSyncLog(t[0], t.slice(1));
    const { name: n } = this.app, i = `"${n.replace(/["`$\\!\n\r;|&<>(){}*?[\]#~]/g, "")} would like to update"`, o = this.sudoWithArgs(i);
    this._logger.info(`Running as non-root user, using sudo to install: ${o}`);
    let a = '"';
    return (/pkexec/i.test(o[0]) || o[0] === "sudo") && (a = ""), this.spawnSyncLog(o[0], [...o.length > 1 ? o.slice(1) : [], `${a}/bin/bash`, "-c", `'${t.join(" ")}'${a}`]);
  }
  sudoWithArgs(t) {
    const n = this.determineSudoCommand(), r = [n];
    return /kdesudo/i.test(n) ? (r.push("--comment", t), r.push("-c")) : /gksudo/i.test(n) ? r.push("--message", t) : /pkexec/i.test(n) && r.push("--disable-internal-agent"), r;
  }
  hasCommand(t) {
    try {
      return this.spawnSyncLog("command", ["-v", t]), !0;
    } catch {
      return !1;
    }
  }
  determineSudoCommand() {
    const t = ["gksudo", "kdesudo", "pkexec", "beesu"];
    for (const n of t)
      if (this.hasCommand(n))
        return n;
    return "sudo";
  }
  /**
   * Detects the package manager to use based on the available commands.
   * Allows overriding the default behavior by setting the ELECTRON_BUILDER_LINUX_PACKAGE_MANAGER environment variable.
   * If the environment variable is set, it will be used directly. (This is useful for testing each package manager logic path.)
   * Otherwise, it checks for the presence of the specified package manager commands in the order provided.
   * @param pms - An array of package manager commands to check for, in priority order.
   * @returns The detected package manager command or "unknown" if none are found.
   */
  detectPackageManager(t) {
    var n;
    let r = t;
    const i = (n = process.env.ELECTRON_BUILDER_LINUX_PACKAGE_MANAGER) === null || n === void 0 ? void 0 : n.trim();
    i && ($v.test(i) ? r = [i] : this._logger.warn(`ELECTRON_BUILDER_LINUX_PACKAGE_MANAGER "${i}" contains unsafe characters. Ignoring override.`));
    for (const s of r)
      if (this.hasCommand(s))
        return s;
    const o = i ? `ELECTRON_BUILDER_LINUX_PACKAGE_MANAGER override "${i}", ` : "", a = t[0];
    return this._logger.warn(`No package manager found in the list: ${o}${t.join(", ")}. Utilizing default: ${a}`), a;
  }
}
_n.LinuxUpdater = Ov;
Object.defineProperty(Yn, "__esModule", { value: !0 });
Yn.DebUpdater = void 0;
const Iv = ce, wl = Rt, Dv = _n;
class fa extends Dv.LinuxUpdater {
  constructor(t, n) {
    super(t, n);
  }
  /*** @private */
  doDownloadUpdate(t) {
    const n = t.updateInfoAndProvider.provider, r = (0, Iv.findFile)(n.resolveFiles(t.updateInfoAndProvider.info), "deb", ["AppImage", "rpm", "pacman"]);
    return this.executeDownload({
      fileExtension: "deb",
      fileInfo: r,
      downloadUpdateOptions: t,
      task: async (i, o) => {
        this.listenerCount(wl.DOWNLOAD_PROGRESS) > 0 && (o.onProgress = (a) => this.emit(wl.DOWNLOAD_PROGRESS, a)), await this.httpExecutor.download(r.url, i, o);
      }
    });
  }
  doInstall(t) {
    const n = this.installerPath;
    if (n == null)
      return this.dispatchError(new Error("No update filepath provided, can't quit and install")), !1;
    if (!this.hasCommand("dpkg") && !this.hasCommand("apt"))
      return this.dispatchError(new Error("Neither dpkg nor apt command found. Cannot install .deb package.")), !1;
    const r = ["dpkg", "apt"], i = this.detectPackageManager(r);
    try {
      fa.installWithCommandRunner(i, n, this.runCommandWithSudoIfNeeded.bind(this), this._logger);
    } catch (o) {
      return this.dispatchError(o), !1;
    }
    return t.isForceRunAfter && this.app.relaunch(), !0;
  }
  static installWithCommandRunner(t, n, r, i) {
    var o;
    if (t === "dpkg")
      try {
        r(["dpkg", "-i", n]);
      } catch (a) {
        i.warn((o = a.message) !== null && o !== void 0 ? o : a), i.warn("dpkg installation failed, trying to fix broken dependencies with apt-get"), r(["apt-get", "install", "-f", "-y"]);
      }
    else if (t === "apt")
      i.warn("Using apt to install a local .deb. This may fail for unsigned packages unless properly configured."), r([
        "apt",
        "install",
        "-y",
        "--allow-unauthenticated",
        // needed for unsigned .debs
        "--allow-downgrades",
        // allow lower version installs
        "--allow-change-held-packages",
        n
      ]);
    else
      throw new Error(`Package manager ${t} not supported`);
  }
}
Yn.DebUpdater = fa;
var Xn = {};
Object.defineProperty(Xn, "__esModule", { value: !0 });
Xn.PacmanUpdater = void 0;
const vl = Rt, Pv = ce, Lv = _n;
class da extends Lv.LinuxUpdater {
  constructor(t, n) {
    super(t, n);
  }
  /*** @private */
  doDownloadUpdate(t) {
    const n = t.updateInfoAndProvider.provider, r = (0, Pv.findFile)(n.resolveFiles(t.updateInfoAndProvider.info), "pacman", ["AppImage", "deb", "rpm"]);
    return this.executeDownload({
      fileExtension: "pacman",
      fileInfo: r,
      downloadUpdateOptions: t,
      task: async (i, o) => {
        this.listenerCount(vl.DOWNLOAD_PROGRESS) > 0 && (o.onProgress = (a) => this.emit(vl.DOWNLOAD_PROGRESS, a)), await this.httpExecutor.download(r.url, i, o);
      }
    });
  }
  doInstall(t) {
    const n = this.installerPath;
    if (n == null)
      return this.dispatchError(new Error("No update filepath provided, can't quit and install")), !1;
    try {
      da.installWithCommandRunner(n, this.runCommandWithSudoIfNeeded.bind(this), this._logger);
    } catch (r) {
      return this.dispatchError(r), !1;
    }
    return t.isForceRunAfter && this.app.relaunch(), !0;
  }
  static installWithCommandRunner(t, n, r) {
    var i;
    try {
      n(["pacman", "-U", "--noconfirm", t]);
    } catch (o) {
      r.warn((i = o.message) !== null && i !== void 0 ? i : o), r.warn("pacman installation failed, attempting to update package database and retry");
      try {
        n(["pacman", "-Sy", "--noconfirm"]), n(["pacman", "-U", "--noconfirm", t]);
      } catch (a) {
        throw r.error("Retry after pacman -Sy failed"), a;
      }
    }
  }
}
Xn.PacmanUpdater = da;
var Kn = {};
Object.defineProperty(Kn, "__esModule", { value: !0 });
Kn.RpmUpdater = void 0;
const Tl = Rt, Fv = ce, xv = _n;
class ha extends xv.LinuxUpdater {
  constructor(t, n) {
    super(t, n);
  }
  /*** @private */
  doDownloadUpdate(t) {
    const n = t.updateInfoAndProvider.provider, r = (0, Fv.findFile)(n.resolveFiles(t.updateInfoAndProvider.info), "rpm", ["AppImage", "deb", "pacman"]);
    return this.executeDownload({
      fileExtension: "rpm",
      fileInfo: r,
      downloadUpdateOptions: t,
      task: async (i, o) => {
        this.listenerCount(Tl.DOWNLOAD_PROGRESS) > 0 && (o.onProgress = (a) => this.emit(Tl.DOWNLOAD_PROGRESS, a)), await this.httpExecutor.download(r.url, i, o);
      }
    });
  }
  doInstall(t) {
    const n = this.installerPath;
    if (n == null)
      return this.dispatchError(new Error("No update filepath provided, can't quit and install")), !1;
    const r = ["zypper", "dnf", "yum", "rpm"], i = this.detectPackageManager(r);
    try {
      ha.installWithCommandRunner(i, n, this.runCommandWithSudoIfNeeded.bind(this), this._logger);
    } catch (o) {
      return this.dispatchError(o), !1;
    }
    return t.isForceRunAfter && this.app.relaunch(), !0;
  }
  static installWithCommandRunner(t, n, r, i) {
    if (t === "zypper")
      return r(["zypper", "--non-interactive", "--no-refresh", "install", "--allow-unsigned-rpm", "-f", n]);
    if (t === "dnf")
      return r(["dnf", "install", "--nogpgcheck", "-y", n]);
    if (t === "yum")
      return r(["yum", "install", "--nogpgcheck", "-y", n]);
    if (t === "rpm")
      return i.warn("Installing with rpm only (no dependency resolution)."), r(["rpm", "-Uvh", "--replacepkgs", "--replacefiles", "--nodeps", n]);
    throw new Error(`Package manager ${t} not supported`);
  }
}
Kn.RpmUpdater = ha;
var Jn = {};
Object.defineProperty(Jn, "__esModule", { value: !0 });
Jn.MacUpdater = void 0;
const Al = he, mo = Nt, Uv = Ct, Sl = ee, kv = ud, Mv = vt, Bv = ce, Cl = ei, bl = er;
class pa extends Mv.AppUpdater {
  constructor(t, n) {
    super(t, n), this.nativeUpdater = Ht.autoUpdater, this.squirrelDownloadedUpdate = !1, this.nativeUpdater.on("error", (r) => {
      this._logger.warn(r), this.emit("error", r);
    }), this.nativeUpdater.on("update-downloaded", () => {
      this.squirrelDownloadedUpdate = !0, this.debug("nativeUpdater.update-downloaded");
    });
  }
  /** Filters update files to the appropriate architecture.
   * On arm64 Macs (including Rosetta), arm64 files are preferred when available.
   * On x64 Macs, arm64 files are excluded. */
  static filterFilesForArch(t, n) {
    const r = (i) => {
      var o;
      return i.url.pathname.includes("arm64") || ((o = i.info.url) === null || o === void 0 ? void 0 : o.includes("arm64"));
    };
    return n && t.some(r) ? t.filter((i) => n === r(i)) : t.filter((i) => !r(i));
  }
  debug(t) {
    this._logger.debug != null && this._logger.debug(t);
  }
  closeServerIfExists() {
    this.server && (this.debug("Closing proxy server"), this.server.close((t) => {
      t && this.debug("proxy server wasn't already open, probably attempted closing again as a safety check before quit");
    }));
  }
  async doDownloadUpdate(t) {
    let n = t.updateInfoAndProvider.provider.resolveFiles(t.updateInfoAndProvider.info);
    const r = this._logger, i = "sysctl.proc_translated";
    let o = !1;
    try {
      this.debug("Checking for macOS Rosetta environment"), o = (0, Cl.execFileSync)("sysctl", [i], { encoding: "utf8" }).includes(`${i}: 1`), r.info(`Checked for macOS Rosetta environment (isRosetta=${o})`);
    } catch (l) {
      r.warn(`sysctl shell command to check for macOS Rosetta environment failed: ${l}`);
    }
    let a = !1;
    try {
      this.debug("Checking for arm64 in uname");
      const f = (0, Cl.execFileSync)("uname", ["-a"], { encoding: "utf8" }).includes("ARM");
      r.info(`Checked 'uname -a': arm64=${f}`), a = a || f;
    } catch (l) {
      r.warn(`uname shell command to check for arm64 failed: ${l}`);
    }
    a = a || process.arch === "arm64" || o, n = pa.filterFilesForArch(n, a);
    const s = (0, Bv.findFile)(n, "zip", ["pkg", "dmg"]);
    if (s == null)
      throw (0, Al.newError)(`ZIP file not provided: ${(0, Al.safeStringifyJson)(n)}`, "ERR_UPDATER_ZIP_FILE_NOT_FOUND");
    const c = t.updateInfoAndProvider.provider, p = "update.zip";
    return this.executeDownload({
      fileExtension: "zip",
      fileInfo: s,
      downloadUpdateOptions: t,
      task: async (l, f) => {
        const h = Sl.join(this.downloadedUpdateHelper.cacheDir, p), g = () => (0, mo.pathExistsSync)(h) ? !t.disableDifferentialDownload : (r.info("Unable to locate previous update.zip for differential download (is this first install?), falling back to full download"), !1);
        let _ = !0;
        g() && (_ = await this.differentialDownloadInstaller(s, t, l, c, p)), _ && await this.httpExecutor.download(s.url, l, f);
      },
      done: async (l) => {
        if (!t.disableDifferentialDownload)
          try {
            const f = Sl.join(this.downloadedUpdateHelper.cacheDir, p);
            await (0, mo.copyFile)(l.downloadedFile, f);
          } catch (f) {
            this._logger.warn(`Unable to copy file for caching for future differential downloads: ${f.message}`);
          }
        return this.updateDownloaded(s, l);
      }
    });
  }
  async updateDownloaded(t, n) {
    var r;
    const i = n.downloadedFile, o = (r = t.info.size) !== null && r !== void 0 ? r : (await (0, mo.stat)(i)).size, a = this._logger, s = `fileToProxy=${t.url.href}`;
    this.closeServerIfExists(), this.debug(`Creating proxy server for native Squirrel.Mac (${s})`), this.server = (0, kv.createServer)(), this.debug(`Proxy server for native Squirrel.Mac is created (${s})`), this.server.on("close", () => {
      a.info(`Proxy server for native Squirrel.Mac is closed (${s})`);
    });
    const c = (p) => {
      const l = p.address();
      return typeof l == "string" ? l : `http://127.0.0.1:${l == null ? void 0 : l.port}`;
    };
    return await new Promise((p, l) => {
      const f = (0, bl.randomBytes)(64).toString("base64").replace(/\//g, "_").replace(/\+/g, "-"), h = Buffer.from(`autoupdater:${f}`, "ascii"), g = `/${(0, bl.randomBytes)(64).toString("hex")}.zip`;
      this.server.on("request", (_, E) => {
        const T = _.url;
        if (a.info(`${T} requested`), T === "/") {
          if (!_.headers.authorization || _.headers.authorization.indexOf("Basic ") === -1) {
            E.statusCode = 401, E.statusMessage = "Invalid Authentication Credentials", E.end(), a.warn("No authenthication info");
            return;
          }
          const P = _.headers.authorization.split(" ")[1], B = Buffer.from(P, "base64").toString("ascii"), [q, Q] = B.split(":");
          if (q !== "autoupdater" || Q !== f) {
            E.statusCode = 401, E.statusMessage = "Invalid Authentication Credentials", E.end(), a.warn("Invalid authenthication credentials");
            return;
          }
          const Z = Buffer.from(`{ "url": "${c(this.server)}${g}" }`);
          E.writeHead(200, { "Content-Type": "application/json", "Content-Length": Z.length }), E.end(Z);
          return;
        }
        if (!T.startsWith(g)) {
          a.warn(`${T} requested, but not supported`), E.writeHead(404), E.end();
          return;
        }
        a.info(`${g} requested by Squirrel.Mac, pipe ${i}`);
        let A = !1;
        E.on("finish", () => {
          A || (this.nativeUpdater.removeListener("error", l), p([]));
        });
        const R = (0, Uv.createReadStream)(i);
        R.on("error", (P) => {
          try {
            E.end();
          } catch (B) {
            a.warn(`cannot end response: ${B}`);
          }
          A = !0, this.nativeUpdater.removeListener("error", l), l(new Error(`Cannot pipe "${i}": ${P}`));
        }), E.writeHead(200, {
          "Content-Type": "application/zip",
          "Content-Length": o
        }), R.pipe(E);
      }), this.debug(`Proxy server for native Squirrel.Mac is starting to listen (${s})`), this.server.listen(0, "127.0.0.1", () => {
        this.debug(`Proxy server for native Squirrel.Mac is listening (address=${c(this.server)}, ${s})`), this.nativeUpdater.setFeedURL({
          url: c(this.server),
          headers: {
            "Cache-Control": "no-cache",
            Authorization: `Basic ${h.toString("base64")}`
          }
        }), this.dispatchUpdateDownloaded(n), this.autoInstallOnAppQuit ? (this.nativeUpdater.once("error", l), this.nativeUpdater.checkForUpdates()) : p([]);
      });
    });
  }
  handleUpdateDownloaded() {
    this.autoRunAppAfterInstall ? this.nativeUpdater.quitAndInstall() : this.app.quit(), this.closeServerIfExists();
  }
  quitAndInstall() {
    this.squirrelDownloadedUpdate ? this.handleUpdateDownloaded() : (this.nativeUpdater.on("update-downloaded", () => this.handleUpdateDownloaded()), this.autoInstallOnAppQuit || this.nativeUpdater.checkForUpdates());
  }
}
Jn.MacUpdater = pa;
var Qn = {}, ma = {};
Object.defineProperty(ma, "__esModule", { value: !0 });
ma.verifySignature = jv;
const Nl = he, af = ei, Hv = ti, Rl = ee;
function sf(e, t) {
  return ['set "PSModulePath=" & chcp 65001 >NUL & powershell.exe', ["-NoProfile", "-NonInteractive", "-InputFormat", "None", "-Command", e], {
    shell: !0,
    timeout: t
  }];
}
function jv(e, t, n) {
  return new Promise((r, i) => {
    const o = t.replace(/'/g, "''");
    n.info(`Verifying signature ${o}`), (0, af.execFile)(...sf(`"Get-AuthenticodeSignature -LiteralPath '${o}' | ConvertTo-Json -Compress"`, 20 * 1e3), (a, s, c) => {
      var p;
      try {
        if (a != null || c) {
          go(n, a, c, i), r(null);
          return;
        }
        const l = Gv(s);
        if (l.Status === 0) {
          try {
            const _ = Rl.normalize(l.Path), E = Rl.normalize(t);
            if (n.info(`LiteralPath: ${_}. Update Path: ${E}`), _ !== E) {
              go(n, new Error(`LiteralPath of ${_} is different than ${E}`), c, i), r(null);
              return;
            }
          } catch (_) {
            n.warn(`Unable to verify LiteralPath of update asset due to missing data.Path. Skipping this step of validation. Message: ${(p = _.message) !== null && p !== void 0 ? p : _.stack}`);
          }
          const h = (0, Nl.parseDn)(l.SignerCertificate.Subject);
          let g = !1;
          for (const _ of e) {
            const E = (0, Nl.parseDn)(_);
            if (E.size ? g = Array.from(E.keys()).every((A) => E.get(A) === h.get(A)) : _ === h.get("CN") && (n.warn(`Signature validated using only CN ${_}. Please add your full Distinguished Name (DN) to publisherNames configuration`), g = !0), g) {
              r(null);
              return;
            }
          }
        }
        const f = `publisherNames: ${e.join(" | ")}, raw info: ` + JSON.stringify(l, (h, g) => h === "RawData" ? void 0 : g, 2);
        n.warn(`Sign verification failed, installer signed with incorrect certificate: ${f}`), r(f);
      } catch (l) {
        go(n, l, null, i), r(null);
        return;
      }
    });
  });
}
function Gv(e) {
  const t = JSON.parse(e);
  delete t.PrivateKey, delete t.IsOSBinary, delete t.SignatureType;
  const n = t.SignerCertificate;
  return n != null && (delete n.Archived, delete n.Extensions, delete n.Handle, delete n.HasPrivateKey, delete n.SubjectName), t;
}
function go(e, t, n, r) {
  if (qv()) {
    e.warn(`Cannot execute Get-AuthenticodeSignature: ${t || n}. Ignoring signature validation due to unsupported powershell version. Please upgrade to powershell 3 or higher.`);
    return;
  }
  try {
    (0, af.execFileSync)(...sf("ConvertTo-Json test", 10 * 1e3));
  } catch (i) {
    e.warn(`Cannot execute ConvertTo-Json: ${i.message}. Ignoring signature validation due to unsupported powershell version. Please upgrade to powershell 3 or higher.`);
    return;
  }
  t != null && r(t), n && r(new Error(`Cannot execute Get-AuthenticodeSignature, stderr: ${n}. Failing signature validation due to unknown stderr.`));
}
function qv() {
  const e = Hv.release();
  return e.startsWith("6.") && !e.startsWith("6.3");
}
Object.defineProperty(Qn, "__esModule", { value: !0 });
Qn.NsisUpdater = void 0;
const Ur = he, $l = ee, Wv = Wt, Vv = fr, Ol = Rt, zv = ce, Yv = Nt, Xv = ma, Il = bt;
class Kv extends Wv.BaseUpdater {
  constructor(t, n) {
    super(t, n), this._verifyUpdateCodeSignature = (r, i) => (0, Xv.verifySignature)(r, i, this._logger);
  }
  /**
   * The verifyUpdateCodeSignature. You can pass [win-verify-signature](https://github.com/beyondkmp/win-verify-trust) or another custom verify function: ` (publisherName: string[], path: string) => Promise<string | null>`.
   * The default verify function uses [windowsExecutableCodeSignatureVerifier](https://github.com/electron-userland/electron-builder/blob/master/packages/electron-updater/src/windowsExecutableCodeSignatureVerifier.ts)
   */
  get verifyUpdateCodeSignature() {
    return this._verifyUpdateCodeSignature;
  }
  set verifyUpdateCodeSignature(t) {
    t && (this._verifyUpdateCodeSignature = t);
  }
  /*** @private */
  doDownloadUpdate(t) {
    const n = t.updateInfoAndProvider.provider, r = (0, zv.findFile)(n.resolveFiles(t.updateInfoAndProvider.info), "exe");
    return this.executeDownload({
      fileExtension: "exe",
      downloadUpdateOptions: t,
      fileInfo: r,
      task: async (i, o, a, s) => {
        const c = r.packageInfo, p = c != null && a != null;
        if (p && t.disableWebInstaller)
          throw (0, Ur.newError)(`Unable to download new version ${t.updateInfoAndProvider.info.version}. Web Installers are disabled`, "ERR_UPDATER_WEB_INSTALLER_DISABLED");
        !p && !t.disableWebInstaller && this._logger.warn("disableWebInstaller is set to false, you should set it to true if you do not plan on using a web installer. This will default to true in a future version."), (p || t.disableDifferentialDownload || await this.differentialDownloadInstaller(r, t, i, n, Ur.CURRENT_APP_INSTALLER_FILE_NAME)) && await this.httpExecutor.download(r.url, i, o);
        const l = await this.verifySignature(i);
        if (l != null)
          throw await s(), (0, Ur.newError)(`New version ${t.updateInfoAndProvider.info.version} is not signed by the application owner: ${l}`, "ERR_UPDATER_INVALID_SIGNATURE");
        if (p && await this.differentialDownloadWebPackage(t, c, a, n))
          try {
            await this.httpExecutor.download(new Il.URL(c.path), a, {
              headers: t.requestHeaders,
              cancellationToken: t.cancellationToken,
              sha512: c.sha512
            });
          } catch (f) {
            try {
              await (0, Yv.unlink)(a);
            } catch {
            }
            throw f;
          }
      }
    });
  }
  // $certificateInfo = (Get-AuthenticodeSignature 'xxx\yyy.exe'
  // | where {$_.Status.Equals([System.Management.Automation.SignatureStatus]::Valid) -and $_.SignerCertificate.Subject.Contains("CN=siemens.com")})
  // | Out-String ; if ($certificateInfo) { exit 0 } else { exit 1 }
  async verifySignature(t) {
    let n;
    try {
      if (n = (await this.configOnDisk.value).publisherName, n == null)
        return null;
    } catch (r) {
      if (r.code === "ENOENT")
        return null;
      throw r;
    }
    return await this._verifyUpdateCodeSignature(Array.isArray(n) ? n : [n], t);
  }
  doInstall(t) {
    const n = this.installerPath;
    if (n == null)
      return this.dispatchError(new Error("No update filepath provided, can't quit and install")), !1;
    const r = ["--updated"];
    t.isSilent && r.push("/S"), t.isForceRunAfter && r.push("--force-run"), this.installDirectory && r.push(`/D=${this.installDirectory}`);
    const i = this.downloadedUpdateHelper == null ? null : this.downloadedUpdateHelper.packageFile;
    i != null && r.push(`--package-file=${i}`);
    const o = () => {
      this.spawnLog($l.join(process.resourcesPath, "elevate.exe"), [n].concat(r)).catch((a) => this.dispatchError(a));
    };
    return t.isAdminRightsRequired ? (this._logger.info("isAdminRightsRequired is set to true, run installer using elevate.exe"), o(), !0) : (this.spawnLog(n, r).catch((a) => {
      const s = a.code;
      this._logger.info(`Cannot run installer: error code: ${s}, error message: "${a.message}", will be executed again using elevate if EACCES, and will try to use electron.shell.openItem if ENOENT`), s === "UNKNOWN" || s === "EACCES" ? o() : s === "ENOENT" ? Ht.shell.openPath(n).catch((c) => this.dispatchError(c)) : this.dispatchError(a);
    }), !0);
  }
  async differentialDownloadWebPackage(t, n, r, i) {
    if (n.blockMapSize == null)
      return !0;
    try {
      const o = {
        newUrl: new Il.URL(n.path),
        oldFile: $l.join(this.downloadedUpdateHelper.cacheDir, Ur.CURRENT_APP_PACKAGE_FILE_NAME),
        logger: this._logger,
        newFile: r,
        requestHeaders: this.requestHeaders,
        isUseMultipleRangeRequest: i.isUseMultipleRangeRequest,
        cancellationToken: t.cancellationToken
      };
      this.listenerCount(Ol.DOWNLOAD_PROGRESS) > 0 && (o.onProgress = (a) => this.emit(Ol.DOWNLOAD_PROGRESS, a)), await new Vv.FileWithEmbeddedBlockMapDifferentialDownloader(n, this.httpExecutor, o).download();
    } catch (o) {
      return this._logger.error(`Cannot download differentially, fallback to full download: ${o.stack || o}`), process.platform === "win32";
    }
    return !1;
  }
}
Qn.NsisUpdater = Kv;
(function(e) {
  var t = Oe && Oe.__createBinding || (Object.create ? function(T, A, R, P) {
    P === void 0 && (P = R);
    var B = Object.getOwnPropertyDescriptor(A, R);
    (!B || ("get" in B ? !A.__esModule : B.writable || B.configurable)) && (B = { enumerable: !0, get: function() {
      return A[R];
    } }), Object.defineProperty(T, P, B);
  } : function(T, A, R, P) {
    P === void 0 && (P = R), T[P] = A[R];
  }), n = Oe && Oe.__exportStar || function(T, A) {
    for (var R in T) R !== "default" && !Object.prototype.hasOwnProperty.call(A, R) && t(A, T, R);
  };
  Object.defineProperty(e, "__esModule", { value: !0 }), e.NsisUpdater = e.MacUpdater = e.RpmUpdater = e.PacmanUpdater = e.DebUpdater = e.AppImageUpdater = e.Provider = e.NoOpLogger = e.AppUpdater = e.BaseUpdater = void 0;
  const r = Nt, i = ee;
  var o = Wt;
  Object.defineProperty(e, "BaseUpdater", { enumerable: !0, get: function() {
    return o.BaseUpdater;
  } });
  var a = vt;
  Object.defineProperty(e, "AppUpdater", { enumerable: !0, get: function() {
    return a.AppUpdater;
  } }), Object.defineProperty(e, "NoOpLogger", { enumerable: !0, get: function() {
    return a.NoOpLogger;
  } });
  var s = ce;
  Object.defineProperty(e, "Provider", { enumerable: !0, get: function() {
    return s.Provider;
  } });
  var c = zn;
  Object.defineProperty(e, "AppImageUpdater", { enumerable: !0, get: function() {
    return c.AppImageUpdater;
  } });
  var p = Yn;
  Object.defineProperty(e, "DebUpdater", { enumerable: !0, get: function() {
    return p.DebUpdater;
  } });
  var l = Xn;
  Object.defineProperty(e, "PacmanUpdater", { enumerable: !0, get: function() {
    return l.PacmanUpdater;
  } });
  var f = Kn;
  Object.defineProperty(e, "RpmUpdater", { enumerable: !0, get: function() {
    return f.RpmUpdater;
  } });
  var h = Jn;
  Object.defineProperty(e, "MacUpdater", { enumerable: !0, get: function() {
    return h.MacUpdater;
  } });
  var g = Qn;
  Object.defineProperty(e, "NsisUpdater", { enumerable: !0, get: function() {
    return g.NsisUpdater;
  } }), n(Rt, e);
  let _;
  function E() {
    if (process.platform === "win32")
      _ = new Qn.NsisUpdater();
    else if (process.platform === "darwin")
      _ = new Jn.MacUpdater();
    else {
      _ = new zn.AppImageUpdater();
      try {
        const T = i.join(process.resourcesPath, "package-type");
        if (!(0, r.existsSync)(T))
          return _;
        switch ((0, r.readFileSync)(T).toString().trim()) {
          case "deb":
            _ = new Yn.DebUpdater();
            break;
          case "rpm":
            _ = new Kn.RpmUpdater();
            break;
          case "pacman":
            _ = new Xn.PacmanUpdater();
            break;
          default:
            break;
        }
      } catch (T) {
        console.warn("Unable to detect 'package-type' for autoUpdater (rpm/deb/pacman support). If you'd like to expand support, please consider contributing to electron-builder", T.message);
      }
    }
    return _;
  }
  Object.defineProperty(e, "autoUpdater", {
    enumerable: !0,
    get: () => _ || E()
  });
})(ze);
function Jv() {
  const e = lt.join($e.getPath("userData"), "updater.log"), t = (f) => {
    try {
      Ce.existsSync(e) && Ce.statSync(e).size > 1024 * 1024 && Ce.renameSync(e, `${e}.previous`), Ce.appendFileSync(e, `${(/* @__PURE__ */ new Date()).toISOString()} ${f}
`);
    } catch (h) {
      console.error("Güncelleme günlüğü yazılamadı:", h);
    }
  }, n = async (f, h = "") => {
    await Pn.showMessageBox({ type: "info", title: "Mülk Yönetim — Güncelleme", message: f, detail: h });
  };
  let r = !1, i, o = !1;
  const a = async () => {
    if (!(o || !i)) {
      o = !0;
      try {
        (await Pn.showMessageBox({
          type: "info",
          title: "Güncelleme hazır",
          message: `${i} sürümü indirildi.`,
          detail: "Açık formlarınızı kaydettikten sonra yeniden başlatabilirsiniz. Ertelerseniz güncelleme uygulamayı kapattığınızda kurulur.",
          buttons: ["Yeniden Başlat ve Kur", "Daha Sonra"],
          defaultId: 1,
          cancelId: 1
        })).response === 0 && ze.autoUpdater.quitAndInstall(!1, !0);
      } finally {
        o = !1;
      }
    }
  }, s = async (f = !1) => {
    if (!$e.isPackaged) {
      f && await n("Güncelleme kontrolü kurulu uygulamada kullanılabilir.", "Geliştirme modunda otomatik güncelleme yapılmaz.");
      return;
    }
    if (i) {
      f && await a();
      return;
    }
    if (r) {
      f && await n("Güncelleme kontrolü veya indirme devam ediyor.");
      return;
    }
    r = !0;
    try {
      t(`Kontrol başladı; mevcut sürüm: ${$e.getVersion()}`);
      const h = await ze.autoUpdater.checkForUpdates();
      h != null && h.downloadPromise ? (f && n(`${h.updateInfo.version} sürümü indiriliyor.`, "İndirme tamamlandığında kurulum seçeneği gösterilecek.").catch((g) => t(String(g))), await h.downloadPromise) : f && await n(`Uygulamanız güncel (${$e.getVersion()}).`);
    } catch (h) {
      t(`Güncelleme başarısız: ${String(h)}`), f && await n("Güncelleme kontrol edilemedi veya indirilemedi.", `İnternet bağlantınızı kontrol edip yeniden deneyin. Sorun sürerse günlük dosyası: ${e}`);
    } finally {
      r = !1;
    }
  };
  if (me.handle("check-for-updates", () => s(!0)), !$e.isPackaged) return;
  ze.autoUpdater.autoDownload = !0, ze.autoUpdater.autoInstallOnAppQuit = !0, ze.autoUpdater.allowPrerelease = !1, ze.autoUpdater.allowDowngrade = !1, ze.autoUpdater.on("error", (f) => t(`Updater: ${String(f)}`)), ze.autoUpdater.on("update-available", (f) => t(`Yeni sürüm: ${f.version}`)), ze.autoUpdater.on("update-not-available", () => t("Uygulama güncel")), ze.autoUpdater.on("update-downloaded", (f) => {
    i = f.version, t(`İndirildi: ${f.version}`), a().catch((h) => t(String(h)));
  });
  const c = () => {
    s().catch((f) => t(String(f)));
  }, p = setTimeout(c, 1e4), l = setInterval(c, 60 * 60 * 1e3);
  $e.once("before-quit", () => {
    clearTimeout(p), clearInterval(l);
  });
}
const Dl = lt.dirname(sd(import.meta.url)), Qv = ld(import.meta.url), Zv = Qv("better-sqlite3"), Zr = $e.getPath("userData");
Ce.existsSync(Zr) || Ce.mkdirSync(Zr, { recursive: !0 });
const Pl = lt.join(Zr, "documents");
Ce.existsSync(Pl) || Ce.mkdirSync(Pl, { recursive: !0 });
const lf = lt.join(Zr, "mulk_yonetim.sqlite");
let G;
try {
  G = new Zv(lf);
} catch (e) {
  console.error("Veritabanı başlatılamadı:", e);
}
G.exec(`
  CREATE TABLE IF NOT EXISTS properties (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT,
    address TEXT,
    status TEXT,
    rent_amount REAL,
    tenant TEXT,
    tenant_phone TEXT,
    tenant_email TEXT,
    payment_status TEXT,
    deduction_amount REAL,
    deduction_note TEXT,
    payment_due_day INTEGER
  );

  CREATE TABLE IF NOT EXISTS payments (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    property_id INTEGER,
    month TEXT,          -- Örn: "Ocak", "Şubat"
    year INTEGER,         -- Örn: 2026
    payment_date TEXT,    -- Örn: "2026-06-24" (Ödeme tarihi)
    status TEXT,          -- Örn: "Ödendi", "Gecikti", "Bekliyor"
    amount_paid REAL,     -- Ödenen miktar (kesintiler düşülmüş net miktar için)
    note TEXT,            -- Ödeme ile ilgili özel notlar
    FOREIGN KEY(property_id) REFERENCES properties(id) ON DELETE CASCADE
  );

  CREATE TABLE IF NOT EXISTS transfers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    source_account TEXT,  -- Örn: "Nakit", "Havale (Tahir Çam)"
    amount REAL,          -- Çıkan miktar
    description TEXT,     -- Nereye/Niçin gönderildiği
    transfer_date TEXT    -- İşlem tarihi
  );

`);
try {
  G.exec("ALTER TABLE payments ADD COLUMN deduction_amount REAL DEFAULT 0");
} catch {
}
try {
  G.exec("ALTER TABLE properties ADD COLUMN payment_due_day INTEGER DEFAULT 1");
} catch {
}
try {
  G.exec("ALTER TABLE properties ADD COLUMN tenant_phone TEXT");
} catch {
}
try {
  G.exec("ALTER TABLE properties ADD COLUMN tenant_email TEXT");
} catch {
}
try {
  G.exec("ALTER TABLE properties ADD COLUMN contract_start_date TEXT");
} catch {
}
try {
  G.exec("ALTER TABLE properties ADD COLUMN contract_end_date TEXT");
} catch {
}
try {
  G.exec("ALTER TABLE properties ADD COLUMN note TEXT");
} catch {
}
try {
  G.exec("ALTER TABLE properties ADD COLUMN contract_file_path TEXT");
} catch {
}
try {
  G.exec("ALTER TABLE properties ADD COLUMN rent_history TEXT DEFAULT '[]'");
} catch {
}
try {
  G.exec("ALTER TABLE payments ADD COLUMN payment_method TEXT DEFAULT 'Nakit'");
} catch {
}
try {
  G.exec("ALTER TABLE payments ADD COLUMN tahir_payi REAL DEFAULT 0");
} catch {
}
try {
  G.exec("ALTER TABLE payments ADD COLUMN bekir_payi REAL DEFAULT 0");
} catch {
}
try {
  G.exec("ALTER TABLE payments ADD COLUMN tarik_payi REAL DEFAULT 0");
} catch {
}
try {
  G.exec("ALTER TABLE payments ADD COLUMN halil_payi REAL DEFAULT 0");
} catch {
}
try {
  G.exec("ALTER TABLE payments ADD COLUMN ozcam_payi REAL DEFAULT 0");
} catch {
}
G.prepare("PRAGMA table_info(payments)").all().some((e) => e.name === "ozcam_insaat_payi") || G.exec("ALTER TABLE payments ADD COLUMN ozcam_insaat_payi REAL DEFAULT 0");
try {
  G.exec('ALTER TABLE transfers ADD COLUMN destination_account TEXT DEFAULT "Dışarıya Harcama / Gider"');
} catch {
}
function Ll() {
  const e = lt.join(Dl, "preload.cjs"), t = new xl({
    title: `Mülk Yönetim — ${$e.getVersion()}`,
    width: 1280,
    height: 800,
    show: !1,
    autoHideMenuBar: !0,
    webPreferences: {
      preload: e,
      // Yeni yolu buraya verdik
      sandbox: !1
    }
  });
  t.on("page-title-updated", (r) => {
    r.preventDefault(), t.setTitle(`Mülk Yönetim — ${$e.getVersion()}`);
  }), t.on("ready-to-show", () => {
    t.show(), t.focus();
  }), t.webContents.setWindowOpenHandler((r) => (Fl.openExternal(r.url), { action: "deny" }));
  const n = process.env.VITE_DEV_SERVER_URL || process.env.ELECTRON_RENDERER_URL;
  n ? t.loadURL(n) : t.loadFile(lt.join(Dl, "../dist/index.html"));
}
$e.whenReady().then(() => {
  me.handle("get-app-version", () => $e.getVersion()), Jv(), me.handle("get-properties", () => G.prepare(`
      WITH LatestPayments AS (
        SELECT property_id, status, note, amount_paid, deduction_amount, month, year,
               ROW_NUMBER() OVER(
                 PARTITION BY property_id 
                 ORDER BY year DESC, 
                 CASE month 
                   WHEN 'Ocak' THEN 1 WHEN 'Şubat' THEN 2 WHEN 'Mart' THEN 3 WHEN 'Nisan' THEN 4 
                   WHEN 'Mayıs' THEN 5 WHEN 'Haziran' THEN 6 WHEN 'Temmuz' THEN 7 WHEN 'Ağustos' THEN 8 
                   WHEN 'Eylül' THEN 9 WHEN 'Ekim' THEN 10 WHEN 'Kasım' THEN 11 WHEN 'Aralık' THEN 12 
                 END DESC
               ) as rn
        FROM payments
      )
      SELECT 
        p.id, p.name, p.address, p.status, p.rent_history, p.rent_amount, p.tenant, p.tenant_phone, p.tenant_email, p.payment_due_day,p.contract_start_date, p.contract_end_date,p.note, p.contract_file_path,
        COALESCE(lp.status, '—') as payment_status,
        COALESCE(lp.note, '') as deduction_note,
        COALESCE(lp.deduction_amount, 0) as deduction_amount,
        lp.amount_paid as last_amount_paid,
        lp.month as last_month,
        lp.year as last_year
      FROM properties p
      LEFT JOIN LatestPayments lp ON p.id = lp.property_id AND lp.rn = 1
    `).all()), me.handle("add-property", (e, t) => {
    const n = La(null, t);
    try {
      const i = G.prepare(`
        INSERT INTO properties (
          name, address, status, rent_amount, tenant, tenant_phone, tenant_email, payment_status, deduction_amount, deduction_note, payment_due_day,contract_start_date, contract_end_date,note, contract_file_path, rent_history
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).run(
        t.name,
        t.address,
        t.status,
        t.rent_amount,
        t.tenant,
        t.tenant_phone,
        t.tenant_email,
        t.payment_status,
        t.deduction_amount,
        t.deduction_note,
        t.payment_due_day,
        t.contract_start_date,
        t.contract_end_date,
        t.note,
        t.contract_file_path,
        n
      );
      return { ...t, id: i.lastInsertRowid, rent_history: n };
    } catch (r) {
      throw console.error("Veritabanına kayıt eklenirken hata oluştu:", r), r;
    }
  }), me.handle("update-property", (e, t) => {
    const n = G.prepare("SELECT * FROM properties WHERE id = ?").get(t.id);
    if (!n) throw new Error("Mülk bulunamadı.");
    const r = La(n, t);
    return G.prepare(`
      UPDATE properties 
      SET name = ?, address = ?, status = ?, rent_amount = ?, tenant = ?, tenant_phone = ?, tenant_email = ?, payment_status = ?, deduction_amount = ?, deduction_note = ?, payment_due_day = ?, contract_start_date = ?, contract_end_date = ?, note = ?, contract_file_path = ?, rent_history = ?
      WHERE id = ?
    `).run(
      t.name,
      t.address,
      t.status,
      t.rent_amount,
      t.tenant,
      t.tenant_phone,
      t.tenant_email,
      t.payment_status,
      t.deduction_amount,
      t.deduction_note,
      t.payment_due_day,
      t.contract_start_date,
      // YENİ
      t.contract_end_date,
      t.note,
      t.contract_file_path,
      r,
      t.id
      // Hangi satırın güncelleneceğini belirten ID
    ), t;
  }), me.handle("get-payments", (e, t) => G.prepare("SELECT * FROM payments WHERE property_id = ? ORDER BY year DESC, month DESC").all(t)), me.handle("select-pdf", async () => {
    const e = await Pn.showOpenDialog({
      // multiSelections parametresi eklendi
      properties: ["openFile", "multiSelections"],
      filters: [{ name: "PDF Dosyaları", extensions: ["pdf"] }]
    });
    if (e.canceled || e.filePaths.length === 0) return null;
    const t = $e.getPath("userData"), n = lt.join(t, "documents");
    Ce.existsSync(n) || Ce.mkdirSync(n, { recursive: !0 });
    const r = [];
    for (const i of e.filePaths) {
      const o = lt.basename(i), a = `${Date.now()}_${o}`, s = lt.join(n, a);
      try {
        Ce.copyFileSync(i, s), console.log("✅ PDF Kopyalandı! Hedef Yol:", s), r.push(s);
      } catch (c) {
        console.error("❌ PDF kopyalanırken hata oluştu:", c);
      }
    }
    return r.length > 0 ? r : null;
  }), me.handle("open-pdf", async (e, t) => t && Ce.existsSync(t) ? (await Fl.openPath(t), !0) : !1), me.handle("save-payment", (e, t) => t.id ? (G.prepare(`
        UPDATE payments 
        SET month = ?, year = ?, payment_date = ?, status = ?, amount_paid = ?, deduction_amount = ?, note = ?, payment_method = ?, tahir_payi = ?, bekir_payi = ?, tarik_payi = ?, halil_payi = ?, ozcam_payi = ?, ozcam_insaat_payi = ?
        WHERE id = ?
      `).run(t.month, t.year, t.payment_date, t.status, t.amount_paid, t.deduction_amount, t.note, t.payment_method, t.tahir_payi, t.bekir_payi, t.tarik_payi, t.halil_payi, t.ozcam_payi, t.ozcam_insaat_payi ?? 0, t.id), t) : { id: G.prepare(`
        INSERT INTO payments (property_id, month, year, payment_date, status, amount_paid, deduction_amount, note, payment_method, tahir_payi, bekir_payi, tarik_payi, halil_payi, ozcam_payi, ozcam_insaat_payi)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
      `).run(t.property_id, t.month, t.year, t.payment_date, t.status, t.amount_paid, t.deduction_amount, t.note, t.payment_method, t.tahir_payi, t.bekir_payi, t.tarik_payi, t.halil_payi, t.ozcam_payi, t.ozcam_insaat_payi ?? 0).lastInsertRowid, ...t }), me.handle("delete-payment", (e, t) => (G.prepare("DELETE FROM payments WHERE id = ?").run(t), !0)), me.handle("export-transactions-excel", async () => {
    let e;
    try {
      const t = await Pn.showSaveDialog({
        title: "Tüm İşlemleri Excel Olarak Kaydet",
        defaultPath: `islemler_${(/* @__PURE__ */ new Date()).toISOString().slice(0, 10)}.xlsx`,
        filters: [{ name: "Excel Çalışma Kitabı", extensions: ["xlsx"] }]
      });
      if (t.canceled || !t.filePath) return { canceled: !0 };
      const { payments: n, transfers: r } = G.transaction(() => ({
        payments: G.prepare(`
        SELECT pay.*, COALESCE(p.name, 'Silinmiş mülk') AS property_name,
               COALESCE(p.tenant, '') AS tenant
        FROM payments pay LEFT JOIN properties p ON p.id = pay.property_id
        ORDER BY pay.payment_date, pay.id
      `).all(),
        transfers: G.prepare("SELECT * FROM transfers ORDER BY transfer_date, id").all()
      }))(), i = gd([
        {
          name: "Ödeme Kayıtları",
          headers: ["Kayıt No", "Mülk No", "Mülk", "Kiracı (Güncel)", "Ay", "Yıl", "Ödeme Tarihi", "Durum", "Hesap / Ödeme Yöntemi", "Kayıtlı Tutar", "Kesinti", "Gerçekleşen Tahsilat", "Not"],
          moneyColumns: [9, 10, 11],
          dateColumns: [6],
          rows: n.map((a) => [
            a.id,
            a.property_id,
            a.property_name,
            a.tenant,
            a.month,
            a.year,
            a.payment_date,
            a.status,
            a.payment_method,
            a.amount_paid ?? 0,
            a.deduction_amount ?? 0,
            a.status === "Ödendi" ? a.amount_paid ?? 0 : 0,
            a.note
          ])
        },
        {
          name: "Transfer ve Giderler",
          headers: ["Kayıt No", "İşlem Tarihi", "İşlem Türü", "Çıkış Hesabı", "Hedef Hesap", "Tutar", "Açıklama"],
          moneyColumns: [5],
          dateColumns: [1],
          rows: r.map((a) => [
            a.id,
            a.transfer_date,
            !a.destination_account || a.destination_account === "Dışarıya Harcama / Gider" ? "Gider" : "Hesaplar Arası Transfer",
            a.source_account,
            a.destination_account || "Dışarıya Harcama / Gider",
            a.amount ?? 0,
            a.description
          ])
        }
      ]), o = t.filePath.toLowerCase().endsWith(".xlsx") ? t.filePath : t.filePath + ".xlsx";
      return e = o + `.${Date.now()}.tmp`, await Ce.promises.writeFile(e, i, { flag: "wx" }), await Ce.promises.rename(e, o), { success: !0, path: o };
    } catch (t) {
      return e && await Ce.promises.unlink(e).catch(() => {
      }), console.error("Excel dışa aktarma hatası:", t), { success: !1, error: "Excel dosyası kaydedilemedi. Dosyanın açık olmadığını ve klasöre yazma izniniz olduğunu kontrol edin." };
    }
  }), me.handle("export-database", async () => {
    const e = await Pn.showSaveDialog({
      title: "Veritabanını Dışa Aktar (Yedekle)",
      defaultPath: "mulk_yonetim_yedek.sqlite",
      filters: [{ name: "SQLite Veritabanı", extensions: ["sqlite"] }]
    });
    if (!e.canceled && e.filePath)
      try {
        return Ce.copyFileSync(lf, e.filePath), { success: !0, path: e.filePath };
      } catch (t) {
        return console.error("Yedekleme sırasında hata:", t), { success: !1 };
      }
    return null;
  }), Ll(), me.handle("delete-property", (e, t) => (G.prepare("DELETE FROM properties WHERE id = ?").run(t), !0)), $e.on("activate", function() {
    xl.getAllWindows().length === 0 && Ll();
  });
});
me.handle("get-finance-summary", () => G.prepare(`
      WITH Incomes AS (
        SELECT 
          COALESCE(SUM(tahir_payi), 0) as t_in,
          COALESCE(SUM(bekir_payi), 0) as b_in,
          COALESCE(SUM(tarik_payi), 0) as tr_in,
          COALESCE(SUM(halil_payi), 0) as h_in,
          COALESCE(SUM(ozcam_payi), 0) as oz_in,
          COALESCE(SUM(ozcam_insaat_payi), 0) as oi_in,
          COALESCE(SUM(CASE WHEN payment_method = 'Nakit' THEN amount_paid ELSE 0 END), 0) as n_in
        FROM payments WHERE status = 'Ödendi'
      ),
      Outgoes AS (
        SELECT
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Tahir Çam)' THEN amount ELSE 0 END), 0) as t_out,
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Bekir Çam)' THEN amount ELSE 0 END), 0) as b_out,
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Tarık Çam)' THEN amount ELSE 0 END), 0) as tr_out,
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Halil Çam)' THEN amount ELSE 0 END), 0) as h_out,
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Özçam Yönetim Kurulu)' THEN amount ELSE 0 END), 0) as oz_out,
          COALESCE(SUM(CASE WHEN source_account = 'Havale (Özçam İnşaat)' THEN amount ELSE 0 END), 0) as oi_out,
          COALESCE(SUM(CASE WHEN source_account = 'Nakit' THEN amount ELSE 0 END), 0) as n_out
        FROM transfers
      ),
      TransferIncomes AS (
        SELECT
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Tahir Çam)' THEN amount ELSE 0 END), 0) as t_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Bekir Çam)' THEN amount ELSE 0 END), 0) as b_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Tarık Çam)' THEN amount ELSE 0 END), 0) as tr_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Halil Çam)' THEN amount ELSE 0 END), 0) as h_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Özçam Yönetim Kurulu)' THEN amount ELSE 0 END), 0) as oz_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Havale (Özçam İnşaat)' THEN amount ELSE 0 END), 0) as oi_ti,
          COALESCE(SUM(CASE WHEN destination_account = 'Nakit' THEN amount ELSE 0 END), 0) as n_ti
        FROM transfers
      )
      SELECT 
        (i.t_in + ti.t_ti - o.t_out) as tahir_total,
        (i.b_in + ti.b_ti - o.b_out) as bekir_total,
        (i.tr_in + ti.tr_ti - o.tr_out) as tarik_total,
        (i.h_in + ti.h_ti - o.h_out) as halil_total,
        (i.oz_in + ti.oz_ti - o.oz_out) as ozcam_total,
        (i.oi_in + ti.oi_ti - o.oi_out) as ozcam_insaat_total,
        (i.n_in + ti.n_ti - o.n_out) as nakit_total
      FROM Incomes i, Outgoes o, TransferIncomes ti;
    `).get());
me.handle("get-payments-by-method", (e, t) => G.prepare(`
      SELECT 
        p.name as property_name, 
        p.tenant, 
        pay.month, 
        pay.year, 
        pay.payment_date, 
        pay.amount_paid, 
        pay.deduction_amount
      FROM payments pay
      JOIN properties p ON pay.property_id = p.id
      WHERE pay.payment_method = ? AND pay.status = 'Ödendi'
      ORDER BY 
        pay.year DESC, 
        CASE pay.month 
          WHEN 'Ocak' THEN 1 WHEN 'Şubat' THEN 2 WHEN 'Mart' THEN 3 WHEN 'Nisan' THEN 4 
          WHEN 'Mayıs' THEN 5 WHEN 'Haziran' THEN 6 WHEN 'Temmuz' THEN 7 WHEN 'Ağustos' THEN 8 
          WHEN 'Eylül' THEN 9 WHEN 'Ekim' THEN 10 WHEN 'Kasım' THEN 11 WHEN 'Aralık' THEN 12 
        END DESC
    `).all(t));
me.handle("add-transfer", (e, t) => ({ id: G.prepare(`
      INSERT INTO transfers (source_account, destination_account, amount, description, transfer_date)
      VALUES (?, ?, ?, ?, ?)
    `).run(t.source_account, t.destination_account, t.amount, t.description, t.transfer_date).lastInsertRowid, ...t }));
me.handle("get-transfers", () => G.prepare("SELECT * FROM transfers ORDER BY transfer_date DESC, id DESC").all());
me.handle("delete-transfer", (e, t) => (G.prepare("DELETE FROM transfers WHERE id = ?").run(t), !0));
$e.on("window-all-closed", () => {
  process.platform !== "darwin" && $e.quit();
});
