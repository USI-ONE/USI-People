/* USI People — Organization Chart & Role-Based Access */

(() => {
  // ── Org Chart Data ──
  // This is the seed data. Once the admin UI is built, this will be
  // loaded from the USI_OrgChart SharePoint list instead.

  const ORG_MEMBERS = [
    // Executives — see everything company-wide
    { email: 'sammy@usicomputer.com', name: 'Sammy Wong', title: 'CEO', managerEmail: null, role: 'executive' },
    { email: 'gary@usicomputer.com', name: 'Gary Gonzalez', title: '', managerEmail: 'sammy@usicomputer.com', role: 'executive' },
    { email: 'johnathank@usicomputer.com', name: 'Johnathan Kendrick', title: '', managerEmail: 'sammy@usicomputer.com', role: 'executive' },
    { email: 'RachelW@usicomputer.com', name: 'Rachel Wong', title: '', managerEmail: 'sammy@usicomputer.com', role: 'executive' },
    { email: 'rachel@usicomputer.com', name: 'Rachel Connett', title: '', managerEmail: 'gary@usicomputer.com', role: 'executive' },
    { email: 'CWall@usicomputer.com', name: 'Chris Wall', title: '', managerEmail: 'johnathank@usicomputer.com', role: 'executive' },
    { email: 'Benson@usicomputer.com', name: 'Benson Chiu', title: '', managerEmail: 'RachelW@usicomputer.com', role: 'executive' },
    { email: 'Aaron@usicomputer.com', name: 'Aaron Kendrick', title: '', managerEmail: 'RachelW@usicomputer.com', role: 'executive' },

    // Managers — see their own + direct/indirect reports only
    { email: 'chriss@usicomputer.com', name: 'Chris Sudweeks', title: '', managerEmail: 'CWall@usicomputer.com', role: 'manager' },
    { email: 'Ray@usicomputer.com', name: 'Ray Leota', title: '', managerEmail: 'CWall@usicomputer.com', role: 'manager' },

    // Employees — see only their own data
    { email: 'Mitch@usicomputer.com', name: 'Mitchell Lindsey', title: '', managerEmail: 'rachel@usicomputer.com', role: 'employee' },
    { email: 'Matt.Dasher@usicomputer.com', name: 'Matt Dasher', title: '', managerEmail: 'rachel@usicomputer.com', role: 'employee' },
    { email: 'Dave@usicomputer.com', name: 'Dave Hancey', title: '', managerEmail: 'rachel@usicomputer.com', role: 'employee' },
    { email: 'mark@usicomputer.com', name: 'Mark Foy', title: '', managerEmail: 'chriss@usicomputer.com', role: 'employee' },
    { email: 'MJones@usicomputer.com', name: 'Marc Jones', title: '', managerEmail: 'chriss@usicomputer.com', role: 'employee' },
    { email: 'colton@usicomputer.com', name: 'Colton Parker', title: '', managerEmail: 'chriss@usicomputer.com', role: 'employee' },
    { email: 'Justin@usicomputer.com', name: 'Justin Silotti', title: '', managerEmail: 'Ray@usicomputer.com', role: 'employee' },
    { email: 'ethan@usicomputer.com', name: 'Ethan Sudweeks', title: '', managerEmail: 'Ray@usicomputer.com', role: 'employee' },
    { email: 'jwan@usicomputer.com', name: 'Jonathan Wan', title: '', managerEmail: 'Benson@usicomputer.com', role: 'employee' },
    { email: 'PeterSmith@usicomputer.com', name: 'Peter Smith', title: '', managerEmail: 'Benson@usicomputer.com', role: 'employee' },
    { email: 'Fernando@usicomputer.com', name: 'Fernando Cardenas', title: '', managerEmail: 'Benson@usicomputer.com', role: 'employee' },
    { email: 'lsudweeks@usicomputer.com', name: 'Lee Sudweeks', title: '', managerEmail: 'Benson@usicomputer.com', role: 'employee' },
    { email: 'Ron@usicomputer.com', name: 'Ron Dickson', title: '', managerEmail: 'Aaron@usicomputer.com', role: 'employee' },
    { email: 'Tui@usicomputer.com', name: 'Tui', title: '', managerEmail: 'Aaron@usicomputer.com', role: 'employee' },
    { email: 'lee@usicomputer.com', name: 'Lee Allred', title: '', managerEmail: 'Aaron@usicomputer.com', role: 'employee' }
  ];

  // ── HR Admin emails (SharePoint site owners) ──
  const HR_ADMINS = [
    'CWall@usicomputer.com',
    'sammy@usicomputer.com'
  ];

  // ── Normalize email for comparison ──
  function _norm(email) { return (email || '').toLowerCase().trim(); }

  // ── Lookup functions ──

  // Find a member by email
  USI.orgFind = function(email) {
    return ORG_MEMBERS.find(m => _norm(m.email) === _norm(email)) || null;
  };

  // ── Departments ──
  const DEPARTMENTS = ['Sales', 'Marketing', 'Operations', 'Professional Services (TOS)', 'Production'];
  USI.getDepartments = function() { return [...DEPARTMENTS]; };

  // Get all org members
  USI.orgAll = function() { return [...ORG_MEMBERS]; };

  // Get display name for an email
  USI.orgName = function(email) {
    const m = USI.orgFind(email);
    return m ? m.name : email;
  };

  // ── Role checks ──

  // Is this person an HR Admin (can see everything)?
  USI.isHRAdmin = function(email) {
    return HR_ADMINS.some(a => _norm(a) === _norm(email));
  };

  // Is this person a manager (has direct reports)?
  USI.isManager = function(email) {
    const norm = _norm(email);
    return ORG_MEMBERS.some(m => _norm(m.managerEmail) === norm);
  };

  // Is this person an executive?
  USI.isExecutive = function(email) {
    const m = USI.orgFind(email);
    return m ? m.role === 'executive' : false;
  };

  // ── Hierarchy queries ──

  // Get direct reports for a manager
  USI.getDirectReports = function(managerEmail) {
    const norm = _norm(managerEmail);
    return ORG_MEMBERS.filter(m => _norm(m.managerEmail) === norm);
  };

  // Get ALL reports (direct + indirect, recursive) for a manager
  USI.getAllReports = function(managerEmail) {
    const norm = _norm(managerEmail);
    const result = [];
    const queue = [norm];
    const visited = new Set();
    while (queue.length > 0) {
      const current = queue.shift();
      if (visited.has(current)) continue;
      visited.add(current);
      const directs = ORG_MEMBERS.filter(m => _norm(m.managerEmail) === current);
      directs.forEach(d => {
        result.push(d);
        queue.push(_norm(d.email));
      });
    }
    return result;
  };

  // Get the manager chain (upward) for a person
  USI.getManagerChain = function(email) {
    const chain = [];
    let current = _norm(email);
    const visited = new Set();
    while (current && !visited.has(current)) {
      visited.add(current);
      const member = ORG_MEMBERS.find(m => _norm(m.email) === current);
      if (!member || !member.managerEmail) break;
      chain.push(member.managerEmail);
      current = _norm(member.managerEmail);
    }
    return chain;
  };

  // Get the direct manager for a person
  USI.getManager = function(email) {
    const m = USI.orgFind(email);
    return m ? USI.orgFind(m.managerEmail) : null;
  };

  // ── Access control ──

  // Can viewer see target's data?
  // Rules: HR Admin sees all, Manager sees their reports (all levels), Employee sees only self
  USI.canView = function(viewerEmail, targetEmail) {
    const vNorm = _norm(viewerEmail);
    const tNorm = _norm(targetEmail);

    // Self
    if (vNorm === tNorm) return true;

    // HR Admin sees all
    if (USI.isHRAdmin(viewerEmail)) return true;

    // Executive sees all
    if (USI.isExecutive(viewerEmail)) return true;

    // Manager sees all reports (recursive)
    const allReports = USI.getAllReports(viewerEmail);
    return allReports.some(r => _norm(r.email) === tNorm);
  };

  // Get list of emails whose data this viewer can access
  USI.getViewableEmails = function(viewerEmail) {
    // HR Admin / Executive: everyone
    if (USI.isHRAdmin(viewerEmail) || USI.isExecutive(viewerEmail)) {
      return ORG_MEMBERS.map(m => m.email);
    }
    // Manager: self + all reports
    if (USI.isManager(viewerEmail)) {
      const reports = USI.getAllReports(viewerEmail);
      return [viewerEmail, ...reports.map(r => r.email)];
    }
    // Employee: self only
    return [viewerEmail];
  };

  // Filter an array of items to only those the viewer can see
  // emailField is the property name containing the target email (e.g., 'EmployeeEmail')
  USI.filterByAccess = function(items, viewerEmail, emailField) {
    const viewable = new Set(USI.getViewableEmails(viewerEmail).map(e => _norm(e)));
    return items.filter(item => {
      const target = item[emailField] || (item.fields && item.fields[emailField]) || '';
      return viewable.has(_norm(target));
    });
  };

  // ── Team selector helper ──
  // Returns the list of people a manager can select from (for 1:1s, team reviews, etc.)
  USI.getTeamForSelector = function(viewerEmail) {
    if (USI.isHRAdmin(viewerEmail) || USI.isExecutive(viewerEmail)) {
      return ORG_MEMBERS.filter(m => _norm(m.email) !== _norm(viewerEmail))
        .sort((a, b) => a.name.localeCompare(b.name));
    }
    if (USI.isManager(viewerEmail)) {
      return USI.getAllReports(viewerEmail)
        .sort((a, b) => a.name.localeCompare(b.name));
    }
    return [];
  };
})();
