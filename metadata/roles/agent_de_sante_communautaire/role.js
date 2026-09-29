/**
 * Data replication logic for the 'agent_de_sante_communautaire' role.
 * These functions determine what data is sent to the user's device during sync.
 * @param {object} app - The application execution context.
 */

/**
 * Determines which organization units this user can see.
 * Can return 'all', 'user_location', or a custom query object.
 */
export function getOrgUnitsForSync(app) {
  return 'all';
}

/**
 * Determines which persons this user can see.
 */
export function getPersonsForSync(app) {
  return 'all';
}

/**
 * Determines which data submissions (reports) this user can see.
 */
export function getReportsForSync(app) {
  return 'all';
}
