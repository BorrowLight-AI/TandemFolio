import { ORDERED_LIST_MARKER_PATTERN, OrderedList, TaskList } from '@tiptap/extension-list'
import { Table } from '@tiptap/extension-table'

const taskListTokenizer = TaskList.config.markdownTokenizer!
const orderedListTokenizer = OrderedList.config.markdownTokenizer!
const tableTokenizer = Table.config.markdownTokenizer!
const taskListItem = /^\s*[-+*]\s+\[([ xX])\]\s+/
const orderedListItem = new RegExp(`^\\s*(?:${ORDERED_LIST_MARKER_PATTERN})[.)]\\s+`)
const tableSeparator = /^[ \t|:]*-[ \t|:-]*$/

function firstLine(source: string): string {
  const end = source.indexOf('\n')
  return end < 0 ? source : source.slice(0, end)
}

function startsWithTable(source: string): boolean {
  const firstEnd = source.indexOf('\n')
  if (firstEnd < 0) return false
  const secondStart = firstEnd + 1
  const secondEnd = source.indexOf('\n', secondStart)
  const header = source.slice(0, firstEnd)
  const separator = source.slice(secondStart, secondEnd < 0 ? source.length : secondEnd)
  return header.includes('|') && separator.includes('|') && tableSeparator.test(separator)
}

/**
 * TipTap's task-list tokenizer splits the complete remaining source before
 * discovering that the current line is not a task item. Marked invokes custom
 * block tokenizers at every cursor, so that fallback becomes quadratic for a
 * long document made from ordinary headings and paragraphs.
 */
export const LinearTaskList = TaskList.extend({
  markdownTokenizer: {
    ...taskListTokenizer,
    tokenize(source, tokens, helpers) {
      if (!taskListItem.test(firstLine(source))) return undefined
      return taskListTokenizer.tokenize(source, tokens, helpers)
    },
  },
})

/**
 * OrderedList has the same fallback shape: it splits all remaining lines before
 * checking whether the first one is an ordered-list item.
 */
export const LinearOrderedList = OrderedList.extend({
  markdownTokenizer: {
    ...orderedListTokenizer,
    tokenize(source, tokens, helpers) {
      if (!orderedListItem.test(firstLine(source))) return undefined
      return orderedListTokenizer.tokenize(source, tokens, helpers)
    },
  },
})

/**
 * The upstream table start probe calls split() on the complete remaining
 * source. Only the first two lines can establish a GFM table candidate.
 */
export const LinearTable = Table.extend({
  markdownTokenizer: {
    ...tableTokenizer,
    start: (source) => (startsWithTable(source) ? 0 : -1),
    tokenize(source, tokens, helpers) {
      if (!startsWithTable(source)) return undefined
      return tableTokenizer.tokenize(source, tokens, helpers)
    },
  },
})
