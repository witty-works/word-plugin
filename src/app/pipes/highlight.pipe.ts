import { Pipe, PipeTransform } from '@angular/core';
import TextUtils from "../utils/text.utils";

@Pipe({
  name: 'highlight'
})
export class HighlightPipe implements PipeTransform {

  transform(context: any, highlight: any): unknown {
    if (!context || !highlight.word) return context;
    const escapedWord = TextUtils.escapeRegExp(highlight.word);
    const re = new RegExp(`(${escapedWord})(?!.*${escapedWord})(?![äöüÄÖÜàéèòìÀÉÈÒÌ\\w])`, 'gu');
    const color = this.getExplanationColor(highlight.details.gravity, highlight.details.subcategory);
    context = context.replace(re, `<span class="highlighted-text highlighted-text--${color}">$1</span>`);
    return context;
}


  getExplanationColor(gravity: number, subcategory: string): string {
    if (subcategory === 'corporate_rules') return 'blue';
    if (!gravity) return 'green';
    else if (gravity < 1.5) return 'red';
    else if (gravity > 2.5) return 'yellow';
    else return 'orange';
  }
}
