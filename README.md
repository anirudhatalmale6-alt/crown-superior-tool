# Crown Superior — the tool's door into the website

`CrownAPI.bas` lets the quoting tool read quotes from crownsuperior.com and write
finished quotes back to it, without opening a browser or filling in a web form.

## Putting it in

1. In the tool, click **Developer**, then **Visual Basic**.
   (Alt+F11 does the same thing, but on many laptops the F keys need **Fn** held
   down, which is why it can look like nothing happens.)
2. **File → Import File…** and pick `CrownAPI.bas`.
3. Close that window.
4. Back in the tool: **Developer → Macros → CrownTestConnection → Run.**

Before step 4, put your key on the `CROWN_KEY` line at the top of the module.
You get it from `https://crownsuperior.com/index.php?cf_action=apikey`, signed in
there as an administrator. The key is a password — anyone who has it can read and
write quotes and policies.

## The three things you can run

| Macro | What it does |
| --- | --- |
| `CrownTestConnection` | Says whether the website is answering and the key is right |
| `CrownLoadQuotes` | Puts the newest quotes on a sheet called **Crown Quotes** |
| `CrownWriteQuoteBack` | Asks for a quote number, company, policy number and amount, and writes them onto that quote |

The first column on the Crown Quotes sheet is the quote number. That is what
`CrownWriteQuoteBack` asks for.

## Underneath

Everything goes to one address — `index.php?cf_action=api` — which answers in
JSON, or in CSV when you ask for `format=csv`. It reaches the quote form (18) and
the policy form (11) and nothing else. `do=` can be `ping`, `list`, `get`,
`update` or `create`.
