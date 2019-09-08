// Copyright (C) 2003, 2004 Daisuke Arai <darai@users.sourceforge.jp>
// Copyright (C) 2005, 2007, 2008, 2010, 2013 panacoran <panacoran@users.sourceforge.jp>
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
// $Id$

using System;
using System.ComponentModel;
using System.Collections.Generic;
using Protra.Lib.Config;
using Protra.Lib.Data;
using Protra.Lib.Lang.Builtins;

namespace PtSim
{
    /// <summary>
    /// システムの成績を計算するクラス
    /// </summary>
    public class Performance
    {
        private readonly string _name;
        private readonly BrandList _brandList;
        private readonly TimeFrame _timeFrame;

        private int _allTrades; // 全トレード数
        private int _winTrades; // 勝ちトレード数
        private int _maxwinCount; // 連勝回数
        private int _maxloseCount; // 連敗回数
        private int _consecutiveWin; // 連勝・連敗中フラグ
        private float _allProfitRatio; // 全トレード平均損益率
        private float _winProfitRatio; // 勝ちトレード平均利率
        private float _winMaxProfitRatio; // 勝ちトレード最大利率
        private float _loseMaxLossRatio; // 負けトレード最大損率
        private float _allTerm; // 全トレード平均期間
        private float _winTerm; // 勝ちトレード平均期間
        private float _budget; // 必要資金
        private float _bookMaxPosition; // 時価の最大ポジション
        private float _marketMaxPosition; // 簿価の最大ポジション
        private float _totalProfit; // 総利益
        private Dictionary<int, float> _allTradesYear;   // 各年のトレード数
        private Dictionary<int, float> _totalProfitYear; // 各年の総利益
        private Dictionary<int, float> _totalWinTradesYear; // 各年の勝率
        private Dictionary<int, float> _allTradesMonth;   // 各月のトレード数
        private Dictionary<int, float> _totalProfitMonth; // 各月の総利益
        private Dictionary<int, float> _totalWinTradesMonth; // 各月の勝率
        private Dictionary<int, float> _totalMaxDrowDownYear; // 各年の最大ドローダウン
        private Dictionary<int, float> _totalMaxMarketDrowDownYear; // 各年の時価最大ドローダウン
        private Dictionary<int, float> _totalWinMaxProfitYear; // 各年の勝ちトレード最大利益
        private Dictionary<int, float> _totalLoseMaxProfitYear; // 各年の負けトレード最大損失
        private Dictionary<int, DateTime> _maxMarketDrowDawnDate; // Max時価DDのトレードの日付
        private Dictionary<int, float> _totalMaxDrowDownMonth; // 各月の最大ドローダウン
        private Dictionary<int, float> _totalWinMaxProfitMonth; // 各月の勝ちトレード最大利益
        private Dictionary<int, float> _totalLoseMaxProfitMonth; // 各月の負けトレード最大損失
        private float _winTotalProfit; // 勝ちトレード総利益
        private float _winMaxProfit; // 勝ちトレード最大利益
        private float _loseMaxLoss; // 負けトレード最大損失
        private int _runningTrades; // 進行中のトレード数
        private float _marketMaxDrowDown; // 時価の最大ドローダウン
        private float _bookMaxDrowDown; // 簿価の最大ドローダウン
        private DateTime _firstTrade; // 最初のトレードの日付
        private DateTime _lastTrade; // 最後のトレードの日付

        /// <summary>
        /// コンストラクタ
        /// </summary>
        /// <param name="name">システムの名前</param>
        /// <param name="timeFrame">タイムフレーム</param>
        /// <param name="brandList">使用する銘柄リスト</param>
        public Performance(string name, BrandList brandList, TimeFrame timeFrame)
        {
            _name = name;
            _brandList = brandList;
            _timeFrame = timeFrame;
            _consecutiveWin = 0;
            _allTradesYear = new Dictionary<int, float>();
            _totalProfitYear = new Dictionary<int, float>();
            _totalWinTradesYear = new Dictionary<int, float>();
            _allTradesMonth = new Dictionary<int, float>();
            _totalProfitMonth = new Dictionary<int, float>();
            _totalWinTradesMonth = new Dictionary<int, float>();             
            _totalMaxDrowDownYear = new Dictionary<int, float>(); 
            _totalMaxMarketDrowDownYear = new Dictionary<int, float>(); 
            _totalWinMaxProfitYear = new Dictionary<int, float>(); 
            _totalLoseMaxProfitYear = new Dictionary<int, float>();
            _maxMarketDrowDawnDate = new Dictionary<int, DateTime>();
            _totalMaxDrowDownMonth = new Dictionary<int, float>(); 
            _totalWinMaxProfitMonth = new Dictionary<int, float>(); 
            _totalLoseMaxProfitMonth = new Dictionary<int, float>();
        }

        /// <summary>
        /// システムの成績を分析する。
        /// </summary>
        /// <param name="worker">Executeを実行するBackgroundWorker</param>
        /// <param name="appendText">TextBoxのAppendText</param>
        /// <returns>利益の集計値のリスト</returns>
        public PricePairList Calculate(BackgroundWorker worker, AppendTextDelegate appendText)
        {
            var profits = AnalyzeLogs(worker);
            if (profits.Count == 0)
                throw new Exception("取引がありません。");
            if (worker.CancellationPending)
                return null;
            _firstTrade = profits.Dates[0];
            _lastTrade = profits.Dates[profits.Count - 1];
            PrintResult(appendText);
            return profits;
        }

        private int T(int a, int b)
        {
            return a * 100 + b;
        }
                    
        private PricePairList AnalyzeLogs(BackgroundWorker worker)
        {
            var profits = new PricePairList();
            var positionValues = new PricePairList();
            var count = 0;
            if (PriceData.MaxDate == DateTime.MinValue)
                throw new Exception("株価データがありません。");
            using (var logData = new LogData(_name, _timeFrame))
            {
                foreach (var code in _brandList.List)
                {
                    if (worker.CancellationPending)
                        return null;
                    var prices = PriceData.GetPrices(code, _timeFrame);
                    if (prices == null)
                        continue;
                    var logs = logData.GetLog(code);
                    var logIndex = 0;
                    var prevClose = 0;
                    var prevHigh = 0;
                    var position = 0;
                    var totalBuy = 0f;
                    var totalSell = 0f;
                    var startDate = DateTime.MinValue;
                    var averagePrice = 0f;
                    foreach (var price in prices)
                    {
                        if (worker.CancellationPending)
                            return null;
                        var dailyProfit = profits[price.Date];
                        var dailyValue = positionValues[price.Date];
                        var close = price.Close;
                        if (position != 0 && close > 0 && prevClose > 0)
                            dailyProfit.AddMarket(position, close - prevClose); // 時価の前日比を加算する。
                        if (close > 0)
                            prevClose = close;
                        if (price.High > 0)
                            prevHigh = price.High;
                        dailyValue.AddMarket(Math.Abs(position), prevHigh);
                        Log log;
                        if (logIndex == logs.Count || logs[logIndex].Date != price.Date || // 売買が発生しない。
                            (log = logs[logIndex++]).Quantity == 0) // 0株の売買は無視する。
                            continue;
                        var prevPosition = position;
                        var consideration = (float)log.Quantity * log.Price;
                        if (log.Order == Order.Buy)
                        {
                            position += log.Quantity;
                            totalBuy += consideration;
                            if (close > 0)
                                dailyProfit.AddMarket(log.Quantity, close - log.Price);
                        }
                        else
                        {
                            position -= log.Quantity;
                            totalSell += consideration;
                            if (close > 0)
                                dailyProfit.AddMarket(log.Quantity, log.Price - close);
                        }
                        var abs = Math.Abs(position);
                        var prevAbs = Math.Abs(prevPosition);
                        if ((float)position * prevPosition > 0) // トレード継続
                        {
                            if (abs < prevAbs)
                            {
                                dailyValue.AddBook(log.Quantity, -averagePrice);
                                dailyProfit.AddBook(log.Quantity, Math.Sign(position) * (log.Price - averagePrice));
                            }
                            else
                            {
                                averagePrice = (prevAbs * averagePrice + consideration) / abs;
                                dailyValue.AddBook(consideration);
                            }
                            continue;
                        }
                        dailyValue.AddBook(prevAbs, -averagePrice);
                        dailyProfit.AddBook(prevPosition, log.Price - averagePrice);
                        if (position == 0) // トレード終了
                            averagePrice = 0;
                        else // ドテンかトレード開始
                        {
                            dailyValue.AddBook(abs, log.Price);
                            averagePrice = log.Price;
                            startDate = log.Date;
                            if (prevPosition == 0) // トレード開始
                                continue;
                        }
                        // ドテンの補正
                        var realTotalBuy = totalBuy;
                        var realTotalSell = totalSell;
                        totalBuy = totalSell = 0;
                        if (position > 0)
                            realTotalBuy -= (totalBuy = (float)position * log.Price);
                        else if (position < 0)
                            realTotalSell -= (totalSell = (float)-position * log.Price);
                        EvaluateTrade(log.Order == Order.Sell, (log.Date - startDate).Days, realTotalBuy, realTotalSell, log.Date, dailyProfit);
                    }
                    if (position != 0)
                        _runningTrades++;
                    worker.ReportProgress(100 * ++count / _brandList.List.Count);
                }
            }
            var realPositionValues = positionValues.BookAccumulatedList;
            CalcMaxPosition(realPositionValues);
            var realProfits = profits.AccumulatedList;
            CalcBudget(realProfits, realPositionValues);
            CalcDrowdown(realProfits);
            return realProfits;
        }

        private void EvaluateTrade(bool isLong, int term, float totalBuy, float totalSell, DateTime date, PricePair dailyprofit)
        {
            var year = date.Year;
            var month = date.Month;
            _allTrades++;
            var ratio = isLong ? totalSell / totalBuy - 1 : 1 - totalBuy / totalSell; // 空売りは売りポジションが分母
            _allProfitRatio += ratio;
            _allTerm += term;
            var profit = totalSell - totalBuy;
            _totalProfit += profit;
            // 年度別のトータル資金
            if (_totalProfitYear.ContainsKey(year) == true)
            {
                _totalProfitYear[year] += profit;
                _totalProfitMonth[T(year, month - 1)] += profit;
                _allTradesYear[year]++;
                _allTradesMonth[T(year, month - 1)]++;
                if (profit >= 0)
                {
                   _totalWinTradesYear[year]++;
                   _totalWinTradesMonth[T(year, month - 1)]++;
                   _totalWinMaxProfitYear[year] += profit;
                   _totalWinMaxProfitMonth[T(year, month - 1)] += profit;
                }
                else
                {
                   _totalLoseMaxProfitYear[year] += profit;
                   _totalLoseMaxProfitMonth[T(year, month - 1)] += profit;
                }
                if (totalSell < totalBuy) // 負け
                {
                    _totalMaxDrowDownYear[year] = Math.Min(_totalMaxDrowDownYear[year], ratio);
                    _totalMaxDrowDownMonth[T(year, month - 1)] = Math.Min(_totalMaxDrowDownMonth[T(year, month - 1)], ratio);
                }
                if (_totalMaxMarketDrowDownYear[year] > dailyprofit.Market)
                {
                    _totalMaxMarketDrowDownYear[year] = dailyprofit.Market;
                    _maxMarketDrowDawnDate[year] = date;
                }
            }
            else
            {
                _totalProfitYear.Add(year, profit);
                for (int i = 0; i < 12; i++)
                {
                    _totalProfitMonth.Add(T(year, i), 0);
                    _allTradesMonth.Add(T(year, i), 0);
                    _totalWinTradesMonth.Add(T(year, i), 0);
                    _totalMaxDrowDownMonth.Add(T(year, i), 0);
                    _totalWinMaxProfitMonth.Add(T(year, i), 0);
                    _totalLoseMaxProfitMonth.Add(T(year, i), 0);
                }
                _totalProfitMonth[T(year, month - 1)] = profit;
                _allTradesYear.Add(year, 1);
                _allTradesMonth[T(year, month - 1)] ++;
                if (profit >= 0)
                {
                   _totalWinTradesYear.Add(year, 1);
                   _totalWinTradesMonth[T(year, month - 1)] = 1;
                   _totalWinMaxProfitYear.Add(year, profit);
                   _totalWinMaxProfitMonth[T(year, month - 1)] = profit;
                   _totalLoseMaxProfitYear.Add(year, 0);
                }
                else
                {
                   _totalWinTradesYear.Add(year, 0);
                   _totalWinMaxProfitYear.Add(year, 0);
                   _totalLoseMaxProfitYear.Add(year, profit);
                   _totalLoseMaxProfitMonth[T(year, month - 1)] = profit;
                }
                if (totalSell < totalBuy) // 負け
                {
                    _totalMaxDrowDownYear.Add(year, ratio);
                    _totalMaxDrowDownMonth[T(year, month - 1)] = ratio;
                    _totalMaxMarketDrowDownYear.Add(year, dailyprofit.Market);
                    _maxMarketDrowDawnDate[year] = date;

                }
                else
                {
                    _totalMaxDrowDownYear.Add(year, 0);
                    _totalMaxMarketDrowDownYear.Add(year, 0);
                    _maxMarketDrowDawnDate[year] = date;
                }
            }
            if (totalSell > totalBuy) // 勝ち
            {
                _winTrades++;
                if (_consecutiveWin >= 0)
                {
                     _consecutiveWin++;
                }
                else
                {
                     _consecutiveWin = 1;
                }
                _maxwinCount = Math.Max(_maxwinCount, _consecutiveWin);
                _winProfitRatio += ratio;
                _winMaxProfitRatio = Math.Max(_winMaxProfitRatio, ratio);
                _winMaxProfit = Math.Max(_winMaxProfit, profit);
                _winTotalProfit += profit;
                _winTerm += term;
            }
            else // 負け
            {
                if (0 > _consecutiveWin)
                {
                     _consecutiveWin--;
                }
                else
                {
                     _consecutiveWin = -1;
                }
                _maxloseCount = Math.Max(_maxloseCount, -1 * _consecutiveWin);
                _loseMaxLossRatio = Math.Min(_loseMaxLossRatio, ratio);
                _loseMaxLoss = Math.Min(_loseMaxLoss, profit);
            }
        }

        private void CalcMaxPosition(PricePairList positionValues)
        {
            var max = new PricePair();
            foreach (var value in positionValues)
                max = PricePair.Max(max, value);
            _marketMaxPosition = max.Market;
            _bookMaxPosition = max.Book;
        }

        private void CalcBudget(PricePairList profits, PricePairList positionValues)
        {
            foreach (var date in profits.Dates)
                _budget = Math.Max(_budget, positionValues[date].Book - profits[date].Book);
        }

// ReSharper disable ParameterTypeCanBeEnumerable.Local
        private void CalcDrowdown(PricePairList profits)
// ReSharper restore ParameterTypeCanBeEnumerable.Local
        {
            var maxProfit = new PricePair();
            var maxDrowdown = new PricePair();
            foreach (var profit in profits)
            {
                maxProfit = PricePair.Max(maxProfit, profit);
                maxDrowdown = PricePair.Min(maxDrowdown, profit - maxProfit);
            }
            _marketMaxDrowDown = maxDrowdown.Market;
            _bookMaxDrowDown = maxDrowdown.Book;
        }

        private void PrintResult(AppendTextDelegate appendText)
        {
            var loseTrades = _allTrades - _winTrades;
            var loseProfitRatio = _allProfitRatio - _winProfitRatio;
            var loseTerm = _allTerm - _winTerm;
            var loseTotalLoss = _totalProfit - _winTotalProfit;
            appendText(string.Format(
                "ファイル: {0}\n" +
                "株価データ: {1}\n" +
                "銘柄リスト: {2}\n" +
                "{3:d}～{4:d}における成績です。\n" +
                "----------------------------------------\n" +
                "全トレード数\t\t{5:d}\n" +
                "勝ちトレード数(勝率)\t{6:d}({7:p})\n" +
                "負けトレード数(負率)\t{8:d}({9:p})\n" +
                "\n" +
                "全トレード平均利率\t{10:p}\n" +
                "勝ちトレード平均利率\t{11:p}\n" +
                "負けトレード平均損率\t{12:p}\n" +
                "\n" +
                "勝ちトレード最大利率\t{13:p}\n" +
                "負けトレード最大損率\t{14:p}\n" +
                "\n" +
                "全トレード平均期間\t{15:n}\n" +
                "勝ちトレード平均期間\t{16:n}\n" +
                "負けトレード平均期間\t{17:n}\n" +
                "----------------------------------------\n" +
                "必要資金\t\t{18:c}\n" +
                "最大ポジション(簿価)\t{19:c}\n" +
                "最大ポジション(時価)\t{20:c}\n" +
                "\n" +
                "純利益\t\t\t{21:c}\n" +
                "勝ちトレード総利益\t\t{22:c}\n" +
                "負けトレード総損失\t\t{23:c}\n" +
                "\n" +
                "全トレード平均利益\t{24:c}\n" +
                "勝ちトレード平均利益\t{25:c}\n" +
                "負けトレード平均損失\t{26:c}\n" +
                "\n" +
                "勝ちトレード最大利益\t{27:c}\n" +
                "負けトレード最大損失\t{28:c}\n" +
                "\n" +
                "プロフィットファクター\t\t{29:n}\n" +
                "最大ドローダウン(簿価)\t{30:c}\n" +
                "最大ドローダウン(時価)\t{31:c}\n" +
                "----------------------------------------\n" +
                "現在進行中のトレード数\t{32:d}\n",
                _name,
                _timeFrame == TimeFrame.Daily ? "日足" : "週足",
                _brandList.Name,
                _firstTrade, _lastTrade,
                _allTrades, _winTrades, (float)_winTrades / _allTrades, loseTrades, (float)loseTrades / _allTrades,
                _allProfitRatio / _allTrades, _winProfitRatio / _winTrades, loseProfitRatio / loseTrades,
                _winMaxProfitRatio, _loseMaxLossRatio,
                _allTerm / _allTrades, _winTerm / _winTrades, loseTerm / loseTrades,
                _budget, _bookMaxPosition, _marketMaxPosition,
                _totalProfit, _winTotalProfit, loseTotalLoss,
                _totalProfit / _allTrades, _winTotalProfit / _winTrades, loseTotalLoss / loseTrades,
                _winMaxProfit, _loseMaxLoss,
                _winTotalProfit / -loseTotalLoss,
                _bookMaxDrowDown, _marketMaxDrowDown,
                _runningTrades));
            // Dictionaryの内容をコピーして、List<KeyValuePair<int, float>>に変換
            List<KeyValuePair<int, float>> list = new List<KeyValuePair<int, float>>(_allTradesYear);
            list.Sort((a, b) => b.Key - a.Key);  //降順
            var j = 0;
            var _total = 0;
            var _ave = 5; // 直近○年の平均年利を求める
            foreach (KeyValuePair<int, float> pair in list)
            {
                j += 1;
                _total += (int)_totalProfitYear[pair.Key];
                if (j >= _ave)
                {
                     break;
                }
            }
            appendText(string.Format(
                           "----------------------------------------\n" +
                           "平均年利\t\t{0:p}\n" +
                           "平均年利(直近{1:d}年)\t{2:p}\n" +
                           "最大連勝\t\t{3:d}回\n" +
                           "最大連敗\t\t{4:d}回\n",
                           (_totalProfit / _allTradesYear.Count) / _budget,
                           _ave,
                           (_total / _ave) / _budget,
                           _maxwinCount,
                           _maxloseCount
                           ));
            appendText(string.Format(
                           "----------------------------------------\n" +
                           "[年度別レポート]\n"));
#if true
            appendText(string.Format("年度\t取引回数\t運用損益\t勝率\tPF\t最大DD\t最大DD(時価)\n"));
            foreach (KeyValuePair<int, float> pair in list)
            {
                appendText(string.Format(
                    "{0}年\t{1,5}回\t\t{2:c}円\t{3,6:p}\t{4,5:n}倍\t{5,6:p}\t{6:c}円({7})\n",
                        pair.Key, pair.Value, _totalProfitYear[pair.Key],
                        _totalWinTradesYear[pair.Key] / pair.Value,
                        Math.Abs(_totalWinMaxProfitYear[pair.Key] / _totalLoseMaxProfitYear[pair.Key]),
                        _totalMaxDrowDownYear[pair.Key], _totalMaxMarketDrowDownYear[pair.Key],
                        _maxMarketDrowDawnDate[pair.Key].ToString("yyyy/MM/dd")
                    ));
            }
#else
            appendText(string.Format("年度\t取引回数\t運用損益\t年利\t勝率\tPF\t最大DD\n"));
            foreach (KeyValuePair<int, float> pair in list)
            {
                appendText(string.Format(
                    "{0}年\t{1,5}回\t\t{2:c}円\t{3:p}\t{4:p}\t{5,5:n}倍\t{6,6:p}\n",
                        pair.Key, pair.Value, _totalProfitYear[pair.Key],
                        _totalProfitYear[pair.Key] / _budget,
                        _totalWinTradesYear[pair.Key] / pair.Value,
                        Math.Abs(_totalWinMaxProfitYear[pair.Key] / _totalLoseMaxProfitYear[pair.Key]),
                        _totalMaxDrowDownYear[pair.Key]
                    ));
            }
#endif            
            appendText(string.Format(
                           "----------------------------------------\n" +
                           "[月別レポート]\n"));
            foreach (KeyValuePair<int, float> pair in list)
            {
                appendText(string.Format("[{0}年]\n", pair.Key));
                appendText(string.Format("月度\t取引回数\t運用損益\t勝率\tPF\t最大DD\n"));
                for (int i = 12 - 1; i >= 0; i--)
                {
                    appendText(string.Format(
                                   "   {0,2}月\t{1,5}回\t\t{2:c}円\t{3,6:p}\t{4,5:n}倍\t{5,6:p}\n",
                                   i + 1, _allTradesMonth[T(pair.Key, i)], _totalProfitMonth[T(pair.Key, i)],
                                   _totalWinTradesMonth[T(pair.Key, i)] / _allTradesMonth[T(pair.Key, i)],
                                   Math.Abs(_totalWinMaxProfitMonth[T(pair.Key, i)] / _totalLoseMaxProfitMonth[T(pair.Key, i)]),
                                   _totalMaxDrowDownMonth[T(pair.Key, i)]
                                   ));
                }
            }
        }
    }
}
