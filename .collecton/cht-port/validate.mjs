const MCP = process.env.COLLECTON_MCP_DIST ?? '/Users/gilbertagbodamakou/Documents/shellit/Dev/collecton/collecton-mcp/dist';
const { loadProject } = await import(`${MCP}/project.js`);
const read = await import(`${MCP}/ops/read.js`);
const r = read.validateProject(loadProject(process.argv[2]));
console.log(JSON.stringify({ valid: r.valid, errors: r.errors.slice(0, 15), count: r.errors.length, warnings: r.warnings.length }, null, 1));
