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

The first time it runs it asks for your key and remembers it. **There is nothing
to edit and nothing to paste into the code.** Save the workbook once afterwards
so it keeps the key.

Your key comes from `https://crownsuperior.com/index.php?cf_action=apikey`, signed
in there as an administrator. To change it later, run the **CrownSetKey** macro.
The key is a password — anyone who has it can read and write quotes and policies.

## The things you can run

| Macro | What it does |
| --- | --- |
| `CrownTestConnection` | Says whether the website is answering and the key is right |
| `CrownLoadQuotes` | Puts the newest quotes on a sheet called **Crown Quotes** |
| `CrownWriteQuoteBack` | Asks for a quote number, company, policy number and amount, and writes them onto that quote |
| `CrownSetKey` | Type in a different key |

The first column on the Crown Quotes sheet is the quote number. That is what
`CrownWriteQuoteBack` asks for.

## Underneath

Everything goes to one address — `index.php?cf_action=api` — which answers in
JSON, or in CSV when you ask for `format=csv`. It reaches the quote form (18) and
the policy form (11) and nothing else. `do=` can be `ping`, `list`, `get`,
`update` or `create`.
