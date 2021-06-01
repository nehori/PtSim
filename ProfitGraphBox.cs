// Copyright (C) 2013 panacoran <panacoran@users.sourceforge.jp>
// Copyright (C) 2003 Daisuke Arai <darai@users.sourceforge.jp>
// 
// This program is part of Protra.
//
// Protra is free software: you can redistribute it and/or modify it
// under the terms of the GNU General Public License as published by
// the Free Software Foundation, either version 3 of the License, or
// (at your option) any later version.
//
// This program is distributed in the hope that it will be useful,
// but WITHOUT ANY WARRANTY; without even the implied warranty of
// MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
// GNU General Public License for more details.
//
// You should have received a copy of the GNU General Public License
// along with this program; if not, see <http://www.gnu.org/licenses/>.
// 
// $Id: ProfitGraphBox.cs 491 2013-08-26 08:27:16Z panacoran $

using System;
using System.Drawing;
using System.Windows.Forms;
using Protra.Lib;

namespace PtSim.Controls
{
    /// <summary>
    /// 利益曲線を描くためのコントロール
    /// </summary>
    public class ProfitGraphBox : UserControl
    {
        private PricePairList _profitList;
        private Graphics _graphics;
        private readonly Color _marketColor = Color.Red;
        private readonly Color _bookColor = Color.Green;
        private readonly Pen _tickPen = new Pen(Color.FromArgb(0x80, Color.Gray));
        private float _priceLabelHeight;
        private RectangleF _chartRectangle;
        private float _upperBound;
        private float _lowerBound;
        private float _xscale;

        /// <summary>
        /// 描画する利益の集計値のリストを取得または設定する。
        /// </summary>
        public PricePairList ProfitList
        {
            get { return _profitList; }
            set
            {
                _profitList = value;
                Invalidate();
            }
        }

        /// <summary>
        /// Double Bufferingを有効にする。
        /// </summary>
        /// <param name="e">イベントの引数</param>
        protected override void OnLoad(EventArgs e)
        {
            DoubleBuffered = true;
        }

        /// <summary>
        /// ペイントイベントを処理する。
        /// </summary>
        /// <param name="e">イベントの引数</param>
        protected override void OnPaint(PaintEventArgs e)
        {
            if (ProfitList == null || ProfitList.Count < 2)
                return;
            _graphics = e.Graphics;
            _priceLabelHeight = _graphics.MeasureString("1", Font).Height;
            _chartRectangle = new RectangleF(ClientRectangle.X, ClientRectangle.Y + _priceLabelHeight,
                                             ClientRectangle.Width, ClientRectangle.Height - _priceLabelHeight * 2);
            _xscale = (_chartRectangle.Width - 2) / ProfitList.Count; // 一番右まで描くとわかりにくいので2px空ける。
            DrawVirticalScale();
            DrawHorizontalScale();
            DrawProfitLines();
            DrawLegends();
        }

        private void DrawVirticalScale()
        {
            var max = float.MinValue;
            var min = float.MaxValue;
            foreach (var profit in ProfitList)
            {
                max = Math.Max(max, Math.Max(profit.Market, profit.Book));
                min = Math.Min(min, Math.Min(profit.Market, profit.Book));
            }
            if (FloatComparers.Equal(max, min))
            {
                max++;
                min--;
            }
            var ticks = (int)(_chartRectangle.Height / (_priceLabelHeight * 4));
            if (ticks < 1)
                ticks = 1;
            var tickRange = TickRange(min, max, ticks);
            _upperBound = (float)(Math.Ceiling(max / tickRange) * tickRange);
            _lowerBound = (float)(Math.Floor(min / tickRange) * tickRange);
            var realTicks = (_upperBound - _lowerBound) / tickRange;
            var tick = _lowerBound;
            var y = _chartRectangle.Bottom;
            var tickSize = _chartRectangle.Height / realTicks;
            while (FloatComparers.Compare(tick, _upperBound) <= 0)
            {
                _graphics.DrawLine(_tickPen, _chartRectangle.Left, y, _chartRectangle.Right, y);
                _graphics.DrawString(tick.ToString("n0"), Font, new SolidBrush(ForeColor),
                                     _chartRectangle.Left, y - _priceLabelHeight);
                tick += tickRange;
                y -= tickSize;
            }
        }

        private static float TickRange(float min, float max, int ticks)
        {
            var raw = (max - min) / ticks;
            var base10 = Math.Floor(Math.Log10(raw));
            if (base10 < 0) // 1より小さい目盛りは作らない。
                base10 = 0;
            var power10 = Math.Pow(10, base10);
            var topdigit = Math.Ceiling(raw / power10);
            return (float)(topdigit * power10);
        }

        private void DrawHorizontalScale()
        {
            int years = 0;

            int startIndex = 0;
            int endIndex = ProfitList.Count - 1;

            // 最初の取引を探す
            for (int i = 0; i < ProfitList.Count; i++)
            {
                if (ProfitList[i].Market != 0.0)
                {
                    if ((ProfitList.Dates[i].Year == ProfitList.Dates[startIndex].Year) && (ProfitList.Dates[i].Month == ProfitList.Dates[startIndex].Month))
                    {
                        // 最初の取引が全データの最初の年月なので開始インデックスは全データの先頭のままとする
                        break;
                    }

                    // 表示上の見易さの為に最初の取引の当年の開始日を開始インデックスにする
                    for (int j = 0; j < ProfitList.Dates.Count; j++)
                    {
                        // 1月の場合は前年の12月
                        if (ProfitList.Dates[i].Month == 1)
                        {
                            if ((ProfitList.Dates[j].Year == (ProfitList.Dates[i].Year - 1)) && (ProfitList.Dates[j].Month == 12))
                            {
                                startIndex = j;
                                break;
                            }
                        }
                        else
                        {
                            if (ProfitList.Dates[j].Year == ProfitList.Dates[i].Year)
                            {
                                startIndex = j;
                                break;
                            }
                        }
                    }

                    break;
                }
            }

            // 最後の取引を探す
            for (int i = endIndex; i > 0; i--)
            {
                if (ProfitList[i].Book != ProfitList[(i - 1)].Book)
                {
                    if ((ProfitList.Dates[i].Year == ProfitList.Dates[endIndex].Year) && (ProfitList.Dates[i].Month == ProfitList.Dates[endIndex].Month))
                    {
                        // 最後の取引が全データの最後の年月なので終了インデックスは全データの末端のままとする
                        break;
                    }

                    // 表示上の見易さの為に最後の取引の翌月の終了日を終了インデックスにする
                    for (int j = endIndex; i > 0; j--)
                    {
                        // 12月の場合は翌年の1月
                        if (ProfitList.Dates[i].Month == 12)
                        {
                            if ((ProfitList.Dates[j].Year == (ProfitList.Dates[i].Year + 1)) && (ProfitList.Dates[j].Month == 1))
                            {
                                endIndex = j;
                                break;
                            }
                        }
                        else
                        {
                            if ((ProfitList.Dates[j].Year == ProfitList.Dates[i].Year) && (ProfitList.Dates[j].Month == (ProfitList.Dates[i].Month + 1)))
                            {
                                endIndex = j;
                                break;
                            }
                        }
                    }

                    break;
                }
            }

            //_xscale = (_chartRectangle.Width - 2) / (ProfitList.Dates.Count - startIndex);
            _xscale = (_chartRectangle.Width - 2) / (endIndex - startIndex);

            years = ProfitList.Dates[ProfitList.Dates.Count - 1].Year - ProfitList.Dates[startIndex].Year + 1;

            var tickRange = 1;
            var labelWidth = _graphics.MeasureString("99", Font).Width;
            while (_chartRectangle.Width * tickRange / years < labelWidth * 1.1)
                tickRange++;
            var smallTickRange = tickRange == 1 ? 3 : 6;
            var prevMonth = -1;

            int startYear = ProfitList.Dates[startIndex].Year;

            const int tickHeight = 3;

            //for (int i = startIndex; i < ProfitList.Dates.Count; i++)
            for (int i = startIndex; i <= endIndex; i++)
            {
                var date = ProfitList.Dates[i];
                if (date.Month == prevMonth)
                    continue;
                prevMonth = date.Month;
                float x = _xscale * (i - startIndex);
                var bottom = _chartRectangle.Bottom;
                if (date.Month == 1)
                {
                    _graphics.DrawLine(_tickPen, x, bottom + tickHeight, x, 0);
                    if ((date.Year - startYear) % tickRange == 0 && x + labelWidth < _chartRectangle.Right)
                        _graphics.DrawString(date.ToString("yy"), Font, new SolidBrush(ForeColor),
                                             x, bottom + tickHeight);
                }
                else if ((date.Month - 1) % smallTickRange == 0)
                    _graphics.DrawLine(_tickPen, x, bottom + tickHeight, x, bottom - tickHeight);
            }
        }

        private void DrawProfitLines()
        {
            int startIndex = 0;
            int endIndex = ProfitList.Count - 1;

            // 最初の取引を探す
            for (int i = 0; i < ProfitList.Count; i++)
            {
                if (ProfitList[i].Market != 0.0)
                {
                    if ((ProfitList.Dates[i].Year == ProfitList.Dates[startIndex].Year) && (ProfitList.Dates[i].Month == ProfitList.Dates[startIndex].Month))
                    {
                        // 最初の取引が全データの最初の年月なので開始インデックスは全データの先頭のままとする
                        break;
                    }

                    // 表示上の見易さの為に最初の取引の当年の開始日を開始インデックスにする
                    for (int j = 0; j < ProfitList.Dates.Count; j++)
                    {
                        // 1月の場合は前年の12月
                        if (ProfitList.Dates[i].Month == 1)
                        {
                            if ((ProfitList.Dates[j].Year == (ProfitList.Dates[i].Year - 1)) && (ProfitList.Dates[j].Month == 12))
                            {
                                startIndex = j;
                                break;
                            }
                        }
                        else
                        {
                            if (ProfitList.Dates[j].Year == ProfitList.Dates[i].Year)
                            {
                                startIndex = j;
                                break;
                            }
                        }
                    }

                    break;
                }
            }

            // 最後の取引を探す
            for (int i = endIndex; i > 0; i--)
            {
                if (ProfitList[i].Book != ProfitList[(i - 1)].Book)
                {
                    if ((ProfitList.Dates[i].Year == ProfitList.Dates[endIndex].Year) && (ProfitList.Dates[i].Month == ProfitList.Dates[endIndex].Month))
                    {
                        // 最後の取引が全データの最後の年月なので終了インデックスは全データの末端のままとする
                        break;
                    }

                    // 表示上の見易さの為に最後の取引の翌月の終了日を終了インデックスにする
                    for (int j = endIndex; i > 0; j--)
                    {
                        // 12月の場合は翌年の1月
                        if (ProfitList.Dates[i].Month == 12)
                        {
                            if ((ProfitList.Dates[j].Year == (ProfitList.Dates[i].Year + 1)) && (ProfitList.Dates[j].Month == 1))
                            {
                                endIndex = j;
                                break;
                            }
                        }
                        else
                        {
                            if ((ProfitList.Dates[j].Year == ProfitList.Dates[i].Year) && (ProfitList.Dates[j].Month == (ProfitList.Dates[i].Month + 1)))
                            {
                                endIndex = j;
                                break;
                            }
                        }
                    }

                    break;
                }
            }

            //_xscale = (_chartRectangle.Width - 2) / (ProfitList.Dates.Count - startIndex);
            _xscale = (_chartRectangle.Width - 2) / (endIndex - startIndex);

            var yscale = _chartRectangle.Height / (_upperBound - _lowerBound);
            var marketPrevPoint = new PointF();
            var bookPrevPoint = new PointF();

            //for (int i = startIndex; i < ProfitList.Count; i++)
            for (int i = startIndex; i <= endIndex; i++)
            {
                var profit = ProfitList[i];
                var top = _chartRectangle.Top;
                float x = _xscale * (i - startIndex);
                var marketPoint = new PointF(x, yscale * (_upperBound - profit.Market) + top);
                var bookPoint = new PointF(x, yscale * (_upperBound - profit.Book) + top);
                //if (i > 0)
                if (i > startIndex)
                {
                    _graphics.DrawLine(new Pen(_marketColor), marketPrevPoint, marketPoint);
                    _graphics.DrawLine(new Pen(_bookColor), bookPrevPoint, bookPoint);
                }
                marketPrevPoint = marketPoint;
                bookPrevPoint = bookPoint;
            }
        }

        private void DrawLegends()
        {
            var label = "時価";
            var size = _graphics.MeasureString(label, Font);
            var rectangle = new RectangleF(_chartRectangle.Right - size.Width - 10,
                                           _chartRectangle.Bottom - size.Height - 10,
                                           size.Height, size.Height);
            _graphics.FillRectangle(new SolidBrush(_marketColor), rectangle);
            _graphics.DrawRectangles(new Pen(ForeColor), new[] {rectangle});
            _graphics.DrawString(label, Font, new SolidBrush(ForeColor), rectangle.Right, rectangle.Top);
            label = "簿価";
            size = _graphics.MeasureString(label, Font);
            rectangle.X -= size.Width + size.Height;
            _graphics.FillRectangle(new SolidBrush(_bookColor), rectangle);
            _graphics.DrawRectangles(new Pen(ForeColor), new[] {rectangle});
            _graphics.DrawString(label, Font, new SolidBrush(ForeColor), rectangle.Right, rectangle.Top);
        }
    }
}