import type { InputHTMLAttributes, ReactElement } from 'react'

export interface ColorSwatch {
  name: string
  hex: string
}

export const THEME_COLORS: readonly ColorSwatch[] = [
  { name: 'White', hex: 'FFFFFF' },
  { name: 'Black', hex: '000000' },
  { name: 'Light Gray', hex: 'E7E6E6' },
  { name: 'Blue Gray', hex: '0E2841' },
  { name: 'Blue', hex: '156082' },
  { name: 'Orange', hex: 'E97132' },
  { name: 'Green', hex: '196B24' },
  { name: 'Sky Blue', hex: '0F9ED5' },
  { name: 'Purple', hex: 'A02B93' },
  { name: 'Light Green', hex: '4EA72E' },
]

export const THEME_COLOR_SHADES: readonly (readonly string[])[] = [
  [
    'F2F2F2',
    '7F7F7F',
    'D0CECE',
    'DDEBF7',
    'DDEBF7',
    'FCE4D6',
    'E2F0D9',
    'DDEBF7',
    'E4DFEC',
    'E2F0D9',
  ],
  [
    'D9D9D9',
    '595959',
    'AEAAAA',
    'BDD7EE',
    '9DC3E6',
    'F8CBAD',
    'C6E0B4',
    '9DC3E6',
    'D9E1F2',
    'C6E0B4',
  ],
  [
    'BFBFBF',
    '3F3F3F',
    '757171',
    '8EA9DB',
    '5B9BD5',
    'F4B084',
    'A9D18E',
    '5B9BD5',
    'B4C6E7',
    'A9D18E',
  ],
  [
    'A6A6A6',
    '262626',
    '3A3838',
    '4472C4',
    '2E75B6',
    'C65911',
    '70AD47',
    '00B0F0',
    '8064A2',
    '70AD47',
  ],
  [
    '808080',
    '0D0D0D',
    '171616',
    '203864',
    '1F4E78',
    '843C0C',
    '375623',
    '0070C0',
    '5B315E',
    '385723',
  ],
]

export const STANDARD_COLORS: readonly ColorSwatch[] = [
  { name: 'Dark Red', hex: 'C00000' },
  { name: 'Red', hex: 'FF0000' },
  { name: 'Orange', hex: 'FFC000' },
  { name: 'Yellow', hex: 'FFFF00' },
  { name: 'Light Green', hex: '92D050' },
  { name: 'Green', hex: '00B050' },
  { name: 'Light Blue', hex: '00B0F0' },
  { name: 'Blue', hex: '0070C0' },
  { name: 'Dark Blue', hex: '002060' },
  { name: 'Purple', hex: '7030A0' },
]

export interface ColorPickerStrings {
  themeColors: string
  standardColors: string
  recentColors?: string | undefined
  moreColors?: string | undefined
  auto?: string | undefined
  shadeTip?: ((row: number, column: number) => string) | undefined
  colorName?: ((swatch: ColorSwatch) => string) | undefined
}

export interface ColorPickerProps {
  value?: string | null | undefined
  strings: ColorPickerStrings
  className?: string | undefined
  recentColors?: readonly string[] | undefined
  onPick: (hex: string | null) => void
  moreInputProps?: InputHTMLAttributes<HTMLInputElement> | undefined
}

const normalizeHex = (hex: string): string => `#${hex.replace(/^#/, '').toUpperCase()}`

export function ColorPicker({
  value,
  strings,
  className,
  recentColors,
  onPick,
  moreInputProps,
}: ColorPickerProps): ReactElement {
  const current = value ? normalizeHex(value) : null
  const isSelected = (hex: string): boolean => current === `#${hex}`
  const swatch = (hex: string, title: string, key?: string): ReactElement => (
    <button
      key={key ?? hex}
      type="button"
      className={`gcp-swatch ${isSelected(hex) ? 'selected' : ''}`}
      aria-label={title}
      aria-pressed={isSelected(hex)}
      title={title}
      style={{ background: `#${hex}` }}
      onMouseDown={(event) => event.preventDefault()}
      onClick={() => onPick(`#${hex}`)}
    />
  )

  return (
    <div
      className={`gcp-palette${className ? ` ${className}` : ''}`}
      role="group"
      aria-label={strings.themeColors}
    >
      {strings.auto && (
        <button
          type="button"
          className={`gcp-auto ${!current ? 'selected' : ''}`}
          aria-pressed={!current}
          onMouseDown={(event) => event.preventDefault()}
          onClick={() => onPick(null)}
        >
          {strings.auto}
        </button>
      )}
      <div className="gcp-section-title">{strings.themeColors}</div>
      <div className="gcp-theme-base">
        {THEME_COLORS.map((color) => swatch(color.hex, strings.colorName?.(color) ?? color.name))}
      </div>
      <div className="gcp-theme-shades">
        {THEME_COLOR_SHADES.flatMap((row, rowIndex) =>
          row.map((hex, columnIndex) =>
            swatch(
              hex,
              strings.shadeTip?.(rowIndex + 1, columnIndex + 1) ?? `#${hex}`,
              `${rowIndex}-${columnIndex}-${hex}`,
            ),
          ),
        )}
      </div>
      <div className="gcp-section-title">{strings.standardColors}</div>
      <div className="gcp-standard-row">
        {STANDARD_COLORS.map((color) =>
          swatch(color.hex, strings.colorName?.(color) ?? color.name),
        )}
      </div>
      {strings.recentColors && recentColors && recentColors.length > 0 && (
        <>
          <div className="gcp-section-title">{strings.recentColors}</div>
          <div className="gcp-standard-row">
            {recentColors.map((hex, index) => {
              const bare = hex.replace(/^#/, '').toUpperCase()
              return swatch(bare, `#${bare}`, `recent-${index}-${bare}`)
            })}
          </div>
        </>
      )}
      {strings.moreColors && (
        <label className="gcp-more">
          <span className="gcp-more-icon" aria-hidden="true">
            <svg
              width="16"
              height="16"
              viewBox="0 0 16 16"
              fill="none"
              stroke="currentColor"
              strokeWidth="1.2"
            >
              <path d="M8 12.98a4.98 4.98 0 1 1 4.98-4.98c0 2.44-1.74 2.49-2.74 2.49-.8 0-1.25.5-1.25 1.25 0 .7-.45 1.25-1 1.25Z" />
              <circle cx="8.83" cy="4.93" r="0.71" fill="currentColor" stroke="none" />
              <circle cx="11.07" cy="6.71" r="0.71" fill="currentColor" stroke="none" />
              <circle cx="6.09" cy="5.51" r="0.71" fill="currentColor" stroke="none" />
              <circle cx="4.93" cy="8.25" r="0.71" fill="currentColor" stroke="none" />
            </svg>
          </span>
          {strings.moreColors}
          <input
            type="color"
            aria-label={strings.moreColors}
            value={(current ?? '#4472C4').toLowerCase()}
            onChange={(event) => onPick(normalizeHex(event.currentTarget.value))}
            {...moreInputProps}
          />
        </label>
      )}
    </div>
  )
}
