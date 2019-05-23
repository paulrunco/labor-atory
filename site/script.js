$(document).ready(function(){
    var ctx = $('#labor_chart');
    var barHeight = 80;
    let scatterChart = new Chart(ctx, {
        type: 'line',
            data: {
                labels: ['', 'Chris', 'Jeremy', 'Mike', 'Walter', 'Colin', ''],
                datasets: [
                {
                    label: 'Scatter Dataset',
                    backgroundColor: "yellow",
                    borderColor: "red",
                    fill: false,
                    borderWidth : barHeight,
                    pointRadius : 0,
                    data: [
                        {
                            x: 0,
                            y: 'Chris'
                        }, {
                            x: 3,
                            y: 'Chris'
                        }
                    ]
                },
                {
                    backgroundColor: "rgba(208,255,154,1)",
                    borderColor: "rgba(208,255,154,1)",
                    fill: false,
                    borderWidth : barHeight,
                    pointRadius : 0,
                    data: [
                        {
                            x: 3,
                            y: 'Walter'
                        }, {
                            x: 5,
                            y: 'Walter'
                        }
                    ]
                },
                {

                    label: 'Scatter Dataset',
                    backgroundColor: "rgba(246,156,85,1)",
                    borderColor: "rgba(246,156,85,1)",
                    fill: false,
                    borderWidth : barHeight,
                    pointRadius : 0,
                    data: [
                        {
                            x: 5,
                            y: 'Mike'
                        }, {
                            x: 10,
                            y: 'Mike'
                        }
                    ]
                },
                {
                    backgroundColor: "rgba(208,255,154,1)",
                    borderColor: "rgba(208,255,154,1)",
                    fill: false,
                    borderWidth : barHeight,
                    pointRadius : 0,
                    data: [
                        {
                            x: 10,
                            y: 'Jeremy'
                        }, {
                            x: 13,
                            y: 'Jeremy'
                        }
                    ]
                }
                ]
            },
            options: {
                responsive: true,
                legend : {
                    display : false
                },
                scales: {
                    xAxes: [{
                        type: 'linear',
                        position: 'bottom',
                        ticks : {
                            beginAtzero :true,
                            stepSize : 1
                        }
                    }],
                    yAxes : [{
                        type: 'category',
                    }]
                }
            }
        });






});
