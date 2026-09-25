# Neo4j Aura skills — vendored

Source: https://github.com/neo4j-contrib/neo4j-skills (MIT, see `NEO4J-SKILLS-LICENSE`)
Pinned commit: a678fef3e47ad3bfd3e96eff9163049c3e8a6ff7 (2026-09-22)

| Skill | Version | Purpose |
|---|---|---|
| `neo4j-aura-provisioning-skill` | 1.0.4 | Create/pause/resume/resize/delete AuraDB instances (aura-cli, REST, Terraform) |
| `neo4j-aura-agent-skill` | 1.0.3 | Manage and invoke Aura Agents via the v2beta1 REST API |
| `neo4j-aura-graph-analytics-skill` | 1.0.10 | Serverless Aura Graph Analytics (GDS Sessions) |

Required environment variables (never commit these):
- `AURA_CLIENT_ID`, `AURA_CLIENT_SECRET` — Aura API client credentials
- `AURA_ORG_ID`, `AURA_PROJECT_ID` — for Aura Agent calls
- `NEO4J_URI`, `NEO4J_USERNAME`, `NEO4J_PASSWORD`, `NEO4J_DATABASE` — database connection

To update: re-copy the three skill directories from a newer commit of the source repo and bump the pinned commit above.
