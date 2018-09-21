package xstyle

import (
	"encoding/json"

	"github.com/360EntSecGroup-Skylar/excelize"
)

type TypeKey string

const (
	LEFT   TypeKey = "left"
	RIGHT  TypeKey = "right"
	TOP    TypeKey = "top"
	BOTTOM TypeKey = "bottom"
)

type StyleKey int

const (
	JISSEN StyleKey = 1
	TENSEN StyleKey = 7
)

type Border struct {
	Type  TypeKey  `json:"type,omitempty"`
	Color string   `json:"color,omitempty"`
	Style StyleKey `json:"style,omitempty"`
}

func NewBorder(top, right, bottom, left StyleKey) []Border {
	return []Border{
		{Type: LEFT, Color: "#000000", Style: left},
		{Type: RIGHT, Color: "#000000", Style: right},
		{Type: TOP, Color: "#000000", Style: top},
		{Type: BOTTOM, Color: "#000000", Style: bottom},
	}
}

type Font struct {
	Size   float64 `json:"size,omitempty"`
	Bold   bool    `json:"bold,omitempty"`
	Family string  `json:"family,omitempty"`
}

type Fill struct {
	Type    string   `json:"type,omitempty"`
	Color   []string `json:"color,omitempty"`
	Pattern int      `json:"pattern,omitempty"`
}

type Alignment struct {
	Rotation float64 `json:"text_rotation,ommitempty"`
}

type Style struct {
	Font      *Font      `json:"font,omitempty"`
	Borders   []Border   `json:"border,omitempty"`
	Fill      *Fill      `json:"fill,omitempty"`
	Alignment *Alignment `json:"alignment,omitempty"`
}

func (this *Style) Json() ([]byte, error) {
	return json.Marshal(this)
}

func (this *Style) Compile(xlsx *excelize.File) (int, error) {
	jsonBin, err := this.Json()
	if err != nil {
		return 0, err
	}
	jsonStr := string(jsonBin)

	style, err := xlsx.NewStyle(jsonStr)
	if err != nil {
		return 0, err
	}
	return style, nil
}
