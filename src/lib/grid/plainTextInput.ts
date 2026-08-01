/** Keep debate shorthand exactly as the debater typed it. */
export function disableTextAssistance(input: HTMLTextAreaElement): void {
    input.setAttribute("autocorrect", "off");
    input.setAttribute("autocapitalize", "off");
    input.setAttribute("autocomplete", "off");
    input.spellcheck = false;
}
