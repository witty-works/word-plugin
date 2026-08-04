# Witty Word Plugin

The Microsoft Word add-in from [Witty Works](https://witty.works), our solution
for "operationalizing" inclusive and consistent language within organizations.

The add-in itself is a client. It sends text to the Witty API for review and
renders the results, so it needs a backend to do anything useful. When this
repository was first published that backend was not available — it now is, along
with the rest of the stack.

## The open source stack

Every component we set out to release is now public:

| Component | Repository | What it does |
| --- | --- | --- |
| Word add-in | this repository | Checks text inside Microsoft Word |
| Browser extension | [browser-extension](https://github.com/witty-works/browser-extension) | The same review flow on the web |
| API / rule engine | [nlp_api](https://github.com/witty-works/nlp_api) | spaCy based rule engine for inclusive language review |
| Rule editor | [rule-editor](https://github.com/witty-works/rule-editor) | Manages the review rules via Django Admin |
| Dashboard | [dashboard](https://github.com/witty-works/dashboard) | Administration UI, accounts and organizations |
| Rule data prep | [nlp_jupyter_notebook](https://github.com/witty-works/nlp_jupyter_notebook) | Pre-processes CSV rule data for the API |
| Release tooling | [make-release](https://github.com/witty-works/make-release) | Version bump, Sentry release and push, used by this repo |

The hosted service these point at by default (`api.witty.works`,
`dashboard.witty.works`) requires a Witty account. To run the stack yourself,
deploy `nlp_api` and `dashboard` and change the `api` and `dashboard` URLs in
`src/environments/`.

Beyond running it, this code is also worth reading if you are building a similar
add-in and want to see how we dealt with:

* triggering checks on text inside Microsoft Word
* allowing interaction with the results along with receiving context information
* replacing content through proposed alternatives.

We would welcome help keeping this going — contributions, and in particular any
offers for hosting. Please contact lukas.smith@witty.works.

## Requirements

Node.js 22.22.3+, 24.15.0+ or 26+ (see the `engines` field in `package.json`).

## Initial Local Setup

- Enable inspector:
  https://learn.microsoft.com/en-us/office/dev/add-ins/testing/debug-office-add-ins-on-ipad-and-mac
- Run `npm run ssl:config`
- Run `npm run build:dev`
- Run `npm run dev-server`
- Run `npm run start:desktop`

More infos:
https://learn.microsoft.com/en-us/office/dev/add-ins/testing/sideload-office-add-ins-for-testing

## Development server

The self-signed SSL certificates used for development expire after 30 days. Run
`npm run ssl:config` to create new ones when needed.

Run `npm run dev-server` for a dev server. Run `npm run start:desktop` to open
Word and sideload the add-in automatically.

To run on the web, get the shareable link for a document you have access to,
then run `npm run start:web -- --document https://thedocumenturl` to open Word
in your default browser and sideload the add-in. If this doesn't work on the
first try, you may need to visit https://localhost:4200 first to let your
browser know to trust the self-signed certificate that is being used.

## Code scaffolding

Run `ng generate component component-name` to generate a new component. You can
also use `ng generate directive|pipe|service|class|guard|interface|enum|module`.

## Build

Run `npm run build:dev` to build the project. The build artifacts will be stored
in the `dist/` directory. Use `npm run build` for a production build.

## Running unit tests

Run `npm run test` to execute the unit tests via
[Karma](https://karma-runner.github.io). Use `npm run test:coverage` to generate
a coverage report in `/coverage`.

## Linting

Run `npm run lint` to check the sources, or `npm run lint:fix` to apply the
fixes ESLint can make automatically.

## Debugging

This project supports debugging using any of the following techniques:

- [Use a browser's developer tools](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-in-office-online)
- [Attach a debugger from the task pane](https://docs.microsoft.com/office/dev/add-ins/testing/attach-debugger-from-task-pane)
- [Use F12 developer tools on Windows 10](https://docs.microsoft.com/office/dev/add-ins/testing/debug-add-ins-using-f12-developer-tools-on-windows-10)
- Remove the sideloaded extension by deleting the relevant XML file in `/Users/username/Library/Containers/com.microsoft.Word/Data/Documents/wef`

## Additional resources

- [Office add-in documentation](https://docs.microsoft.com/office/dev/add-ins/overview/office-add-ins)
- More Office Add-in samples at
  [OfficeDev on Github](https://github.com/officedev)

This project has adopted the
[Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information, see the
[Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any
additional questions or comments.

This add-in is built with Angular. To get more help on the Angular CLI use
`ng help` or go check out the
[Angular CLI Overview and Command Reference](https://angular.dev/cli) page.
