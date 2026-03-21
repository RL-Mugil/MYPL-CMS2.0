    const INTERNAL_ROLES = ['Super Admin', 'Admin', 'Galvanizer', 'Staff', 'Attorney'];
    const CLIENT_SIDE_ROLES = ['Client', 'Client Admin', 'Client Employee', 'Individual Client'];

    function getEffectiveSessionRoles(session) {
      const roles = [];
      if (session && session.role) roles.push(session.role);
      String(session && session.additionalRoles || '')
        .split(',')
        .map((value) => value.trim())
        .filter(Boolean)
        .forEach((role) => {
          if (!roles.includes(role)) roles.push(role);
        });
      return roles;
    }

    function hasInternalAccess(session) {
      return getEffectiveSessionRoles(session).some((role) => INTERNAL_ROLES.includes(role));
    }

    function hasClientAccess(session) {
      return getEffectiveSessionRoles(session).some((role) => CLIENT_SIDE_ROLES.includes(role));
    }

    (async () => {
      const status = document.getElementById('login-status');

      try {
        await Auth.initClerk();
        const session = await Auth.requireAuth([...INTERNAL_ROLES, ...CLIENT_SIDE_ROLES]);

        if (session) {
          status.textContent = 'Verifying account...';
          status.className = 'alert alert-info';

          if (hasInternalAccess(session)) {
            window.location.href = 'dashboard.html';
          } else if (hasClientAccess(session)) {
            await Auth.signOut();
            window.location.href = 'client-login.html';
          } else {
            status.textContent = 'This account is authenticated but has no portal role mapping.';
            status.className = 'alert alert-error';
          }
          return;
        }

        if (Auth.getClerkEmail()) {
          status.textContent = 'Your Clerk login succeeded, but this email is not mapped in USER_ROLES yet.';
          status.className = 'alert alert-error';
        }

        await Auth.mountSignIn('clerk-sign-in');
      } catch (e) {
        API.reportError('index-init', e);
        status.textContent = 'Connection error: ' + e.message;
        status.className = 'alert alert-error';
      }
    })();
  