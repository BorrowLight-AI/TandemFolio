import { createElement } from 'react'
import { renderToStaticMarkup } from 'react-dom/server'
import { describe, expect, it } from 'vitest'

import { ColorPicker, STANDARD_COLORS, THEME_COLORS, THEME_COLOR_SHADES } from '@genoffice/ui'

describe('Office color palette', () => {
  it('renders theme colors, theme shades, standard colors, and the optional actions', () => {
    const markup = renderToStaticMarkup(
      createElement(ColorPicker, {
        value: '#4472c4',
        strings: {
          auto: 'Automatic',
          themeColors: 'Theme Colors',
          standardColors: 'Standard Colors',
          moreColors: 'More Colors…',
        },
        onPick: () => {},
      }),
    )

    expect(THEME_COLORS).toHaveLength(10)
    expect(THEME_COLOR_SHADES).toHaveLength(5)
    expect(THEME_COLOR_SHADES.every((row) => row.length === 10)).toBe(true)
    expect(STANDARD_COLORS).toHaveLength(10)
    expect(markup.match(/gcp-swatch/g)).toHaveLength(70)
    expect(markup).toContain('background:#4472C4')
    expect(markup).toContain('gcp-swatch selected')
    expect(markup).toContain('Automatic')
    expect(markup).toContain('More Colors…')
  })
})
