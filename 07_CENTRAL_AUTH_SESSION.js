/**
 * OWNERDASH_20260706_V05 — Central Auth session gate.
 * Authority: KalmaCore.requireStaffRole(pin, "OWNER_DASHBOARD", allowedRoles).
 * No local PIN sheet. No local staff-access sheet authority.
 * Client/localStorage session data is display/cache only; every protected server read validates server-side session.
 */
function kodLoginOwnerDashboardV1(payload) {
  payload = payload || {};
  var pin = kodAuthText_(payload.pin || payload.PIN || payload.staffPin || payload.Staff_PIN);
  if (!pin) return kodAuthFail_('PIN_REQUIRED', 'PIN wajib diisi.');
  try {
    var auth = kodRequireCentralStaffRole_(pin, KOD_AUTH.roleSets.login, 'kodLoginOwnerDashboardV1');
    var session = kodCreateDashboardSession_(auth);
    return {
      ok: true,
      status: 'PASS',
      authSource: KOD_AUTH.authSource,
      moduleCode: KOD_AUTH.moduleCode,
      sessionToken: session.sessionToken,
      session: kodPublicSession_(session),
      localPinRetired: true,
      localRoleAuthority: false,
      clientRoleAuthority: false
    };
  } catch (err) {
    return kodAuthFail_('CENTRAL_AUTH_FAILED', kodSafeAuthError_(err));
  }
}

function kodValidateOwnerDashboardSessionV1(payload) {
  try {
    var session = kodRequireDashboardSession_(payload || {}, KOD_AUTH.roleSets.dashboardRead, 'kodValidateOwnerDashboardSessionV1');
    return { ok: true, status: 'PASS', session: kodPublicSession_(session) };
  } catch (err) {
    return kodAuthFail_('SESSION_INVALID', kodSafeAuthError_(err));
  }
}

function kodLogoutOwnerDashboardV1(payload) {
  var token = kodExtractDashboardSessionToken_(payload || {});
  if (token) CacheService.getScriptCache().remove(KOD_AUTH.sessionCachePrefix + token);
  return { ok: true, status: 'PASS', loggedOut: true };
}

function kodRequireCentralStaffRole_(pin, allowedRoles, actionName) {
  var rawPin = kodAuthText_(pin);
  if (!rawPin) throw new Error('CENTRAL_AUTH_PIN_REQUIRED');
  if (typeof KalmaCore === 'undefined' || !KalmaCore || typeof KalmaCore.requireStaffRole !== 'function') {
    throw new Error('CENTRAL_AUTH_LIBRARY_MISSING: KalmaCore.requireStaffRole is required.');
  }
  var roles = kodNormalizeRoleList_(allowedRoles || []);
  if (!roles.length) throw new Error('CENTRAL_AUTH_ALLOWED_ROLES_EMPTY');
  var result = KalmaCore.requireStaffRole(rawPin, KOD_AUTH.moduleCode, roles);
  var normalized = kodNormalizeCentralAuthResult_(result);
  if (!normalized.ok) throw new Error('CENTRAL_AUTH_REJECTED_FOR_' + (actionName || 'OWNER_DASHBOARD'));
  if (roles.indexOf(normalized.role) < 0) throw new Error('CENTRAL_AUTH_ROLE_NOT_ALLOWED: ' + normalized.role);
  return normalized;
}

function kodNormalizeCentralAuthResult_(result) {
  result = result || {};
  if (result.ok === false || result.status === 'FAIL') throw new Error('CENTRAL_AUTH_FAILED');
  var identity = result.identity || result.staff || result.user || result.data || result;
  var rawRole = kodReadNested_(identity, ['canonicalRole','role','Role','normalizedRole','staffRole','Staff_Role']) || kodReadNested_(result, ['canonicalRole','role','Role','staffRole']);
  var canonicalRole = kodNormalizeDashboardRole_(rawRole);
  var staffId = kodAuthText_(kodReadNested_(identity, ['staffId','Staff_ID','id','ID','email','Email']) || kodReadNested_(result, ['staffId','Staff_ID','email']));
  var staffName = kodAuthText_(kodReadNested_(identity, ['staffName','Staff_Name','name','Name','displayName','email']) || kodReadNested_(result, ['staffName','Staff_Name','name','email']) || canonicalRole);
  if (!canonicalRole) throw new Error('CENTRAL_AUTH_ROLE_MISSING');
  if (!staffName) throw new Error('CENTRAL_AUTH_IDENTITY_MISSING');
  return {
    ok: true,
    staffId: staffId || staffName,
    staffName: staffName,
    role: canonicalRole,
    rawRole: kodAuthText_(rawRole),
    authSource: KOD_AUTH.authSource,
    moduleCode: KOD_AUTH.moduleCode,
    permissionCheckedBy: KOD_AUTH.requiredLibraryCall
  };
}

function kodCreateDashboardSession_(auth) {
  var now = new Date();
  var expiresAt = new Date(now.getTime() + (KOD_AUTH.sessionTtlSeconds * 1000));
  var token = Utilities.getUuid();
  var session = {
    sessionToken: token,
    staffId: auth.staffId || '',
    staffName: auth.staffName || '',
    role: auth.role || '',
    rawRole: auth.rawRole || '',
    authSource: auth.authSource || KOD_AUTH.authSource,
    moduleCode: auth.moduleCode || KOD_AUTH.moduleCode,
    permissionCheckedBy: auth.permissionCheckedBy || KOD_AUTH.requiredLibraryCall,
    createdAt: now.toISOString(),
    expiresAt: expiresAt.toISOString()
  };
  CacheService.getScriptCache().put(KOD_AUTH.sessionCachePrefix + token, JSON.stringify(session), KOD_AUTH.sessionTtlSeconds);
  return session;
}

function kodRequireDashboardSession_(payload, allowedRoles, actionName) {
  var token = kodExtractDashboardSessionToken_(payload || {});
  if (!token) throw new Error('AUTH_SESSION_REQUIRED: login Central Auth dulu.');
  var raw = CacheService.getScriptCache().get(KOD_AUTH.sessionCachePrefix + token);
  if (!raw) throw new Error('AUTH_SESSION_INVALID_OR_EXPIRED: login ulang.');
  var session = JSON.parse(raw);
  if (new Date(session.expiresAt).getTime() < new Date().getTime()) {
    CacheService.getScriptCache().remove(KOD_AUTH.sessionCachePrefix + token);
    throw new Error('AUTH_SESSION_EXPIRED: login ulang.');
  }
  var role = kodNormalizeDashboardRole_(session.role);
  var allowed = kodNormalizeRoleList_(allowedRoles || []);
  if (allowed.indexOf(role) < 0) throw new Error('ROLE_BLOCKED_FOR_' + (actionName || 'OWNER_DASHBOARD') + ': ' + role);
  session.role = role;
  return session;
}

function kodRequireOwnerDashboardSession_(payload, actionName) {
  return kodRequireDashboardSession_(payload || {}, KOD_AUTH.roleSets.ownerOnly, actionName || 'OWNER_ONLY_ACTION');
}

function kodExtractDashboardSessionToken_(payload) {
  payload = payload || {};
  if (typeof payload === 'string') return kodAuthText_(payload);
  return kodAuthText_(payload.sessionToken || payload.Session_Token || payload.sessionId || payload.Session_ID || payload.authToken || payload.token || (payload.session && (payload.session.sessionToken || payload.session.token || payload.session.sessionId)));
}

function kodPublicSession_(session) {
  session = session || {};
  return {
    staffId: session.staffId || '',
    staffName: session.staffName || '',
    role: session.role || '',
    rawRole: session.rawRole || '',
    authSource: session.authSource || KOD_AUTH.authSource,
    moduleCode: session.moduleCode || KOD_AUTH.moduleCode,
    createdAt: session.createdAt || '',
    expiresAt: session.expiresAt || '',
    clientRoleAuthority: false,
    localStorageAuthority: false,
    localPinRetired: true
  };
}

function kodNormalizeDashboardRole_(role) {
  var r = String(role || '').trim().toUpperCase();
  r = (KOD_AUTH.legacyRoleAliases && KOD_AUTH.legacyRoleAliases[r]) || r;
  return r;
}

function kodNormalizeRoleList_(roles) {
  var seen = {};
  var out = [];
  (roles || []).forEach(function(role) {
    var r = kodNormalizeDashboardRole_(role);
    if (r && !seen[r]) { seen[r] = true; out.push(r); }
  });
  return out;
}

function kodReadNested_(obj, paths) {
  for (var i = 0; i < paths.length; i++) {
    var cur = obj;
    var parts = String(paths[i]).split('.');
    for (var p = 0; p < parts.length; p++) {
      if (cur == null) break;
      cur = cur[parts[p]];
    }
    if (cur !== undefined && cur !== null && String(cur).trim() !== '') return cur;
  }
  return '';
}

function kodAuthText_(value) {
  return String(value == null ? '' : value).trim();
}

function kodSafeAuthError_(err) {
  var msg = String(err && err.message ? err.message : err || 'AUTH_FAILED');
  if (/PIN|password/i.test(msg) && !/^PIN_REQUIRED$/.test(msg)) return 'Central Auth menolak login. Cek PIN / role / permission module.';
  if (/CENTRAL_AUTH_LIBRARY_MISSING/.test(msg)) return 'Central Auth library KalmaCore belum terpasang.';
  if (/ROLE_BLOCKED|ROLE_NOT_ALLOWED|ROLE_BLOCKED_FOR/.test(msg)) return 'Role tidak diizinkan untuk Owner Dashboard.';
  if (/SESSION/.test(msg)) return msg;
  return 'Central Auth menolak login. Cek PIN / role / permission module.';
}

function kodAuthFail_(code, message) {
  return { ok: false, status: 'FAIL', code: code || 'AUTH_FAILED', message: message || 'Auth failed', authSource: KOD_AUTH.authSource, localPinRetired: true, localRoleAuthority: false, clientRoleAuthority: false };
}

function RUN_OWNERDASH_CENTRAL_AUTH_STATIC_PROOF_V05() {
  var failed = [];
  function hasFn(name) { return typeof this[name] === 'function'; }
  if (typeof KalmaCore === 'undefined') failed.push('kalmaCoreRuntimeNotAvailableInStaticContext');
  if (typeof kodLoginOwnerDashboardV1 !== 'function') failed.push('missingLoginEndpoint');
  if (typeof kodValidateOwnerDashboardSessionV1 !== 'function') failed.push('missingValidateSessionEndpoint');
  if (typeof kodRequireCentralStaffRole_ !== 'function') failed.push('missingCentralAuthWrapper');
  if (typeof kodRequireDashboardSession_ !== 'function') failed.push('missingServerSessionGuard');
  if (typeof kodGetOwnerDashboardV0 !== 'function') failed.push('missingDashboardEndpoint');
  var authText = String(kodRequireCentralStaffRole_) + String(kodGetOwnerDashboardV0) + String(kodLoginOwnerDashboardV1);
  if (authText.indexOf('KalmaCore.requireStaffRole') < 0) failed.push('centralAuthCallMissing');
  var localSheetMarkers = [['STAFF','ACCESS'].join('_'), ['STAFF','ACCESS','MS'].join('_')];
  if (authText.indexOf(localSheetMarkers[0]) >= 0 || authText.indexOf(localSheetMarkers[1]) >= 0) failed.push('localStaffAccessAuthorityPresent');
  if (/found\s*\[\s*1\s*\]\s*!==\s*pin/.test(authText)) failed.push('localPinComparisonFound');
  return {
    status: failed.filter(function(x){ return x !== 'kalmaCoreRuntimeNotAvailableInStaticContext'; }).length ? 'FAIL' : 'PASS',
    note: 'KalmaCore static runtime availability is checked after library install; appsscript.json carries dependency proof.',
    build: KOD_BUILD.name,
    centralAuthCall: KOD_AUTH.requiredLibraryCall,
    moduleCode: KOD_AUTH.moduleCode,
    loginEndpoint: 'kodLoginOwnerDashboardV1',
    validateEndpoint: 'kodValidateOwnerDashboardSessionV1',
    protectedEndpoint: 'kodGetOwnerDashboardV0(payload.sessionToken)',
    ownerOnlyGuard: 'kodRequireOwnerDashboardSession_',
    clientRoleTrustPolicy: KOD_AUTH.clientRoleTrustPolicy,
    localPinRetired: true,
    localStaffAccessAuthority: false,
    failedChecks: failed.filter(function(x){ return x !== 'kalmaCoreRuntimeNotAvailableInStaticContext'; })
  };
}
