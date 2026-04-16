### Changelog

All notable changes to this project will be documented in this file. Dates are displayed in UTC.

#### 0.1.0

- chore(m365-agents): complete Task 1 — lint, test, build all green [`7ed8f3b`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/7ed8f3b5c3a8b50f8ca1a597f43f941577c871c1)
- chore: initial scaffold before M0 tasks [`bb6d008`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/bb6d0081c4d1ec2827a155cf29b8df675f0c5e49)
- chore(m365-agents): add vitest, shared/ and tests/ scaffold, runtime deps [`a3ffaf1`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/a3ffaf1363db46efeee37ae92f13a76edfab0f13)
- test(m365-agents): add execute/webhook unit tests for all four nodes [`950c167`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/950c16798f083c9220734022e8414aabb1fb9d0c)
- feat(m365-agents): real M365SendActivity covering all five operations [`e245f59`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/e245f59430089fdf661234489848f9e50db751de)
- feat(m365-agents): standalone JWT validator replicating SDK middleware [`8692e91`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/8692e91071c1dec4935e4a9d886e6ad04ed4c0db)
- feat(m365-agents): envelope helpers for Trigger and builders [`efe30d7`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/efe30d7f2de86e80bc7ef704f82c0c0bba47f5c8)
- feat(m365-agents): real M365AgentTrigger with JWT validation + GET health [`d8b967f`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/d8b967f47d5b26d7f2376035624b4e754b3b73da)
- feat(m365-agents): add shared envelope and domain types [`c6592b8`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/c6592b8c149c16d54c838044dd9903539cbdcb00)
- test(m365-agents): add gated Azure integration tests [`28a66b6`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/28a66b6d6927d980f29d0abc05101a3c565362fb)
- feat(m365-agents): M365CardTemplate with adaptivecards-templating [`23cd264`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/23cd2647e9dc237121c01d0bf3942b374b510bb2)
- feat(m365-agents): buildAuthConfig maps credentials to SDK AuthConfiguration [`8476e87`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/8476e87082e16c3eaea6511fa8d63b9d270abaee)
- feat(m365-agents): bot connector bundle with MSAL + axios + replyInThread [`abb90aa`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/abb90aaabbd3244c1fcbd8dde8c18aee8ef2d0fd)
- feat(m365-agents): M365TextMessage builder node [`9c279e6`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/9c279e61d3886737a4ce92075669e657d9ecd7d7)
- chore(m365-agents): justfile with deploy-dev recipe [`0060ed9`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/0060ed9b72823f857c1f9f5ad607c3444dcb98e1)
- docs(m365-agents): update AGENTS.md layout, add README for M0 [`2ed1e35`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/2ed1e3541305eac63ca821d5818fc473acef6d3c)
- docs(m365-agents): rename npm package to n8n-nodes-agents-sdk [`f4153c7`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/f4153c76e228f518e88397ea2db7139ae715922b)
- refactor(m365-agents): apply code-review polish (6 AUTO-FIX items) [`86ac27b`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/86ac27b5577809cccf18e940d27da1521bac1d2a)
- feat(m365-agents): finalize M365AgentApi credential with pinned tests [`bb50243`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/bb5024365a504b15509fe813541c7d04a3fb97cf)
- fix(m365-agents): use configWithoutCloudSupport, restore SDK Activity import [`404d69a`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/404d69a3c114ed9d1465d20b5d06ea35600d60dd)
- perf(m365-agents): memoize BotConnectorBundle by serviceUrl in M365SendActivity [`97af7a3`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/97af7a3f52164b5e3e44cd5b7bb12176af37c9e9)
- docs(m365-agents): update repo URLs to n8n-nodes-agents-sdk [`b76c4c2`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/b76c4c2cacf275099b39395a4343f73ca340e470)
- chore(m365-agents): trim npm publish bundle [`c277323`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/c277323027e56c96eb50685aeedf20d8caf59e18)
- Release 0.1.1 [`816f358`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/816f358fdbc7c786f011b2582e0462ad5d59ad59)
- chore(m365-agents): register all M0 nodes in package.json [`b9ac365`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/b9ac365494286b67d63aa33bb18cb3d8aa1fe4b2)
- fix(m365-agents): reword M365CardTemplate bindingData description [`52a19d0`](https://github.com/mrkhachaturov/n8n-nodes-agents-sdk.git
/commit/52a19d0f98b0a75a77d5fa6f8e5788889fc547ba)
