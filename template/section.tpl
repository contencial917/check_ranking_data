<section class="page">
    <div class="header">
        <div class="header__inner">
            <h2 class="header__title">順位計測結果</h2>
        </div>
    </div>
    <div class="main">
        <div class="main__inner">
            <div class="main__section keyword">
                <dl class="keyword__list">
                    <dt class="keyword__title">URL:</dt>
                    <dd class="keyword__data">https://{ domain }/</dd>
                    <dt class="keyword__title">検索語:</dt>
                    <dd class="keyword__data">{ keyword }</dd>
                </dl>
                <div class="keyword__ranking">
                    <table class="keyword__table">
                        <thead class="keyword__thead">
                            <tr class="keyword__row">
                                <th class="keyword__th left">調査日</th>
                                <th class="keyword__th right">Google</th>
                            </tr>
                        </thead>
                        <tbody>
                            { ranking_table }
                        </tbody>
                    </table>
                    { out_of_range }
                </div>
            </div>
            <div class="main__section graph">
                <h2 class="graph__title">「{ keyword }」検索順位（{ date }）</h2>
                <div class="graph__container">
                    <canvas id="ex_chart_{ page_no }" class="graph__chart"></canvas>
                </div>
                <p class="graph__google">
                    <span class="graph__picture-1"></span>
                    <span class="graph__picture-2"></span>
                    Google
                </p>
            </div>
        </div>
    </div>
    <div class="footer">
        <div class="footer__inner">
            <p class="footer__page">{ page_no } / { total_pages }</p>
        </div>
    </div>
</section>

<script>
    var ctx_{ page_no } = document.getElementById('ex_chart_{ page_no }');

    var data_{ page_no } = {
        labels: [{ labels }],
        datasets: [{
            label: 'Google',
            data: [{ ranking_data }],
            borderColor: 'rgba(63,128,239,1)',
            backgroundColor: 'rgba(63,128,239,1)',
            lineTension: 0,
            fill: false,
            borderWidth: 2
        }]
    };

    var options_{ page_no } = {
        scales: {
            xAxes: [{
                ticks: {
                    fontSize: 15,
                    fontStyle: "bold"
                }
            }],
            yAxes: [{
                ticks: {
                    fontSize: 15,
                    fontStyle: "bold",
                    reverse: true,
                    min: 1,
                    max: { max_size },
                    stepSize: { step_size },
                    userCallback: function (tick) {
                        return tick.toString();
                    }
                }
            }]
        },
        legend: {
            display: false
        },
        responsive: true,
        maintainAspectRatio: false
    };

    var ex_chart_{ page_no } = new Chart(ctx_{ page_no }, {
        type: 'line',
        data: data_{ page_no },
        options: options_{ page_no }
    });
</script>

<div class="pagebreak"></div>