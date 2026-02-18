/**
 * ============================================================
 * Bridge.gs - Client → Server Communication
 * ============================================================
 * The frontend calls:
 *   google.script.run
 *     .withSuccessHandler(fn)
 *     .withFailureHandler(fn)
 *     .handleClientRequest(action, dataJson)
 *
 * This is the ONLY function the client calls directly.
 * It validates the session, routes the request, and returns
 * a plain JavaScript object (GAS auto-serialises to JSON).
 */

function handleClientRequest(action, dataJson) {
  try {
    // ── 1. Session check ────────────────────────────────────
    var userEmail = Session.getActiveUser().getEmail();
    if (!userEmail) {
      return { success: false, error: 'Not authenticated' };
    }

    var user = getUserByEmail(userEmail);
    if (!user) {
      return { success: false, error: 'User not found in system' };
    }
    if (user.status !== 'Active') {
      return { success: false, error: 'Account is inactive' };
    }

    // ── 2. Parse payload ────────────────────────────────────
    var data = {};
    if (dataJson && dataJson !== '') {
      try { data = JSON.parse(dataJson); } catch(pe) { data = {}; }
    }

    // ── 3. Route ────────────────────────────────────────────
    switch (action) {

      // Dashboard
      case 'getDashboardStats':
        return getDashboardStats(user);

      // Daily Reports
      case 'getDSRList':
        return getDailyReports(user, data);
      case 'submitDSR':
        return submitDailyReport(data, user);

      // Travel Plans
      case 'getTravelList':
        return getTravelPlans(user, data);
      case 'submitTravel':
        return submitTravelPlan(data, user);
      case 'updateTravelStatus':
        return updateTravelStatus(data, user);

      // Leads
      case 'getLeads':
        return getLeads(user, data);
      case 'submitLead':
        return submitLead(data, user);
      case 'updateLeadStatus':
        return updateLeadStatus(data, user);

      // Bookings
      case 'getBookings':
        return getBookings(user, data);
      case 'submitBooking':
        return submitBooking(data, user);

      // Reports
      case 'getMonthlyReport':
        return getMonthlyReport(data, user);
      case 'exportReport':
        return exportReport(data, user);

      // Incentives
      case 'getIncentives':
        return getIncentives(user, data);
      case 'calculateIncentive':
        return calculateIncentive(data, user);

      // Admin
      case 'getUsers':
        return getUsers(user);
      case 'addUser':
        return addUser(data, user);
      case 'updateUser':
        return updateUser(data, user);
      case 'getSystemLogs':
        return getSystemLogs(user);
      case 'getEmailRecipients':
        return getEmailRecipients(user);
      case 'updateEmailRecipient':
        return updateEmailRecipient(data, user);
      
      // Flexible Incentive Tiers
      case 'getIncentiveTiers':
        return getIncentiveTiers();
      case 'updateIncentiveTiers':
        return updateIncentiveTiers(data, user);
      case 'recalculateIncentivesForMonth':
        return recalculateIncentivesForMonth(data, user);
      case 'getIncentiveBreakdownForSales':
        return getIncentiveBreakdownForSales(data.totalSales || 0);
      
      // Enhanced PDF Reports
      case 'generateDetailedIncentivePDF':
        return generateDetailedIncentivePDF(data, user);
      case 'generateManagementReportPDF':
        return generateManagementReportPDF(data, user);

      default:
        return { success: false, error: 'Unknown action: ' + action };
    }

  } catch (err) {
    Logger.log('handleClientRequest ERROR [' + action + ']: ' + err.message);
    try { logActivity('system', 'ERROR', '[' + action + '] ' + err.message, 'Error'); } catch(le){}
    return { success: false, error: err.message };
  }
}

