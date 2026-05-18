const WORKBOOK_FILE = 'Evan_Ye_ESports_Earnings.xlsx';

const state = {
  players: [],
  tournaments: [],
  charts: {},
  currentGame: ''
};

const GAME_COLORS = [
  '#3b82f6', '#60a5fa', '#93c5fd', '#2563eb', '#1d4ed8',
  '#14b8a6', '#2dd4bf', '#5eead4', '#0f766e', '#0d9488'
];

const COUNTRY_COLORS = [
  '#f59e0b', '#fbbf24', '#fcd34d', '#f97316', '#fb923c',
  '#14b8a6', '#2dd4bf', '#5eead4', '#3b82f6', '#60a5fa'
];

document.addEventListener('DOMContentLoaded', init);

async function init() {
  try {
    setStatus('Loading workbook...');

    const { players, tournaments } = await loadWorkbook(WORKBOOK_FILE);
    state.players = players;
    state.tournaments = tournaments;

    renderHero();
    renderTopGamesChart();
    renderTopCountriesChart();
    populateGameSelect();
    renderConcentration();

    const defaultGame = aggregateSum(state.players, 'game', 'winnings')[0]?.label || '';
    if (defaultGame) {
      document.getElementById('game-select').value = defaultGame;
      updateSelectedGame(defaultGame);
    }

    document.getElementById('game-select').addEventListener('change', (event) => {
      updateSelectedGame(event.target.value);
    });

    setStatus('Project loaded.');
    window.setTimeout(() => {
      document.getElementById('status-box').classList.add('hidden');
    }, 1000);
  } catch (error) {
    console.error(error);
    setStatus('Could not load the workbook. Make sure all files are in the same folder and open the page with Live Server.');
  }
}

async function loadWorkbook(fileName) {
  const response = await fetch(fileName);
  if (!response.ok) {
    throw new Error(`Workbook not found: ${fileName}`);
  }

  const buffer = await response.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });

  const playersSheet = workbook.Sheets['Winnings by gamer'];
  const tournamentsSheet = workbook.Sheets['Prize by tournament'];

  const playersRaw = XLSX.utils.sheet_to_json(playersSheet, { defval: '' });
  const tournamentsRaw = XLSX.utils.sheet_to_json(tournamentsSheet, { defval: '' });

  const players = playersRaw
    .map((row) => ({
      country: String(row['Country Name'] || '').trim(),
      gamertag: String(row['Gamertag'] || '').trim(),
      name: String(row['Name'] || '').trim(),
      winnings: Number(row['Total Winnings']) || 0,
      game: String(row['Game'] || '').trim(),
      tournamentResults: Number(row['Tournament Results']) || 0,
      link: String(row['Link'] || '').trim()
    }))
    .filter((row) => row.game && row.country && row.winnings > 0);

  const tournaments = tournamentsRaw
    .map((row) => ({
      tournamentName: String(row['Tournament Name'] || '').trim(),
      prizeMoney: Number(row['Prize Money']) || 0,
      game: String(row['Game'] || '').trim(),
      teams: parseCount(row['Teams']),
      players: parseCount(row['Players'])
    }))
    .filter((row) => row.game && row.tournamentName && row.prizeMoney > 0);

  return { players, tournaments };
}

function parseCount(value) {
  const match = String(value ?? '').match(/\d+/);
  return match ? Number(match[0]) : 0;
}

function setStatus(message) {
  document.getElementById('status-box').textContent = message;
}

function money(value) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    maximumFractionDigits: 0
  }).format(value || 0);
}

function compactMoney(value) {
  return new Intl.NumberFormat('en-US', {
    style: 'currency',
    currency: 'USD',
    notation: 'compact',
    compactDisplay: 'short',
    maximumFractionDigits: 1
  }).format(value || 0);
}

function percent(value) {
  return `${(value * 100).toFixed(1)}%`;
}

function aggregateSum(list, groupKey, valueKey) {
  const grouped = new Map();

  list.forEach((item) => {
    const key = item[groupKey];
    if (!key) {
      return;
    }
    const current = grouped.get(key) || 0;
    grouped.set(key, current + (Number(item[valueKey]) || 0));
  });

  return [...grouped.entries()]
    .map(([label, value]) => ({ label, value }))
    .sort((a, b) => b.value - a.value);
}

function renderHero() {
  const uniqueGames = new Set([...state.players.map((d) => d.game), ...state.tournaments.map((d) => d.game)]).size;
  const uniqueCountries = new Set(state.players.map((d) => d.country)).size;
  const totalPlayerWinnings = state.players.reduce((sum, row) => sum + row.winnings, 0);
  const totalTournamentPrize = state.tournaments.reduce((sum, row) => sum + row.prizeMoney, 0);
  const topGame = aggregateSum(state.players, 'game', 'winnings')[0];
  const share = topGame ? topGame.value / totalPlayerWinnings : 0;

  document.getElementById('hero-stats').innerHTML = `
    <div class="stat-card"><h3>Games in the dataset</h3><p>${uniqueGames}</p></div>
    <div class="stat-card"><h3>Countries represented</h3><p>${uniqueCountries}</p></div>
    <div class="stat-card"><h3>Total player winnings</h3><p>${compactMoney(totalPlayerWinnings)}</p></div>
    <div class="stat-card"><h3>Total tournament prize money</h3><p>${compactMoney(totalTournamentPrize)}</p></div>
  `;

  document.getElementById('hero-note').textContent = topGame
    ? `${topGame.label} alone accounts for ${percent(share)} of all player winnings shown in this dataset.`
    : '';
}

function baseChartOptions(titleText) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    plugins: {
      legend: {
        display: false
      },
      title: {
        display: true,
        text: titleText,
        color: '#1f1f1f',
        font: {
          size: 16,
          weight: '600'
        },
        padding: {
          bottom: 16
        }
      },
      tooltip: {
        callbacks: {
          label(context) {
            const value = context.raw;
            if (typeof value === 'number') {
              return money(value);
            }
            return value;
          }
        }
      }
    },
    scales: {
      x: {
        ticks: {
          color: '#4a4a4a'
        },
        grid: {
          color: '#e5e5df'
        }
      },
      y: {
        ticks: {
          color: '#4a4a4a'
        },
        grid: {
          display: false
        }
      }
    }
  };
}

function createOrUpdateChart(key, context, config) {
  if (state.charts[key]) {
    state.charts[key].destroy();
  }
  state.charts[key] = new Chart(context, config);
}

function renderTopGamesChart() {
  const data = aggregateSum(state.players, 'game', 'winnings').slice(0, 10);
  const context = document.getElementById('gamesChart');

  createOrUpdateChart('gamesChart', context, {
    type: 'bar',
    data: {
      labels: data.map((d) => d.label),
      datasets: [
        {
          data: data.map((d) => d.value),
          backgroundColor: GAME_COLORS,
          borderRadius: 6
        }
      ]
    },
    options: {
      ...baseChartOptions('Top 10 games by total player winnings'),
      indexAxis: 'y',
      scales: {
        x: {
          ticks: {
            callback(value) {
              return compactMoney(value);
            },
            color: '#4a4a4a'
          },
          grid: {
            color: '#e5e5df'
          }
        },
        y: {
          ticks: {
            color: '#4a4a4a'
          },
          grid: {
            display: false
          }
        }
      },
      plugins: {
        ...baseChartOptions('').plugins,
        title: {
          display: true,
          text: 'Top 10 games by total player winnings',
          color: '#1f1f1f',
          font: { size: 16, weight: '600' },
          padding: { bottom: 16 }
        },
        tooltip: {
          callbacks: {
            label(context) {
              return `Total winnings: ${money(context.raw)}`;
            }
          }
        }
      }
    }
  });
}

function renderTopCountriesChart() {
  const data = aggregateSum(state.players, 'country', 'winnings').slice(0, 10);
  const context = document.getElementById('countriesChart');

  createOrUpdateChart('countriesChart', context, {
    type: 'bar',
    data: {
      labels: data.map((d) => d.label),
      datasets: [
        {
          data: data.map((d) => d.value),
          backgroundColor: COUNTRY_COLORS,
          borderRadius: 6
        }
      ]
    },
    options: {
      ...baseChartOptions('Top 10 countries by total player winnings'),
      indexAxis: 'y',
      scales: {
        x: {
          ticks: {
            callback(value) {
              return compactMoney(value);
            },
            color: '#4a4a4a'
          },
          grid: {
            color: '#e5e5df'
          }
        },
        y: {
          ticks: {
            color: '#4a4a4a'
          },
          grid: {
            display: false
          }
        }
      },
      plugins: {
        title: {
          display: true,
          text: 'Top 10 countries by total player winnings',
          color: '#1f1f1f',
          font: { size: 16, weight: '600' },
          padding: { bottom: 16 }
        },
        legend: {
          display: false
        },
        tooltip: {
          callbacks: {
            label(context) {
              return `Total winnings: ${money(context.raw)}`;
            }
          }
        }
      }
    }
  });
}

function populateGameSelect() {
  const select = document.getElementById('game-select');
  const games = aggregateSum(state.players, 'game', 'winnings').map((d) => d.label);

  select.innerHTML = games
    .map((game) => `<option value="${escapeHtml(game)}">${game}</option>`)
    .join('');
}

function updateSelectedGame(game) {
  state.currentGame = game;

  const filteredPlayers = state.players.filter((row) => row.game === game);
  const filteredTournaments = state.tournaments.filter((row) => row.game === game);

  const winningsTotal = filteredPlayers.reduce((sum, row) => sum + row.winnings, 0);
  const prizeTotal = filteredTournaments.reduce((sum, row) => sum + row.prizeMoney, 0);
  const topCountry = aggregateSum(filteredPlayers, 'country', 'winnings')[0];
  const topPlayer = [...filteredPlayers].sort((a, b) => b.winnings - a.winnings)[0];

  document.getElementById('selected-cards').innerHTML = `
    <div class="mini-card"><h3>Players in this game</h3><p>${filteredPlayers.length}</p></div>
    <div class="mini-card"><h3>Tournaments in this game</h3><p>${filteredTournaments.length}</p></div>
    <div class="mini-card"><h3>Total player winnings</h3><p>${compactMoney(winningsTotal)}</p></div>
    <div class="mini-card"><h3>Total tournament prize money</h3><p>${compactMoney(prizeTotal)}</p></div>
  `;

  document.getElementById('selected-game-note').textContent = topCountry && topPlayer
    ? `Within ${game}, ${topCountry.label} has the highest total player winnings in this dataset, and ${topPlayer.gamertag} is the top-earning player listed for the game.`
    : `This section updates to show how ${game} compares within the larger esports landscape.`;

  renderSelectedGameChart(filteredPlayers, game);
  renderTournamentChart(filteredTournaments, game);
}

function renderSelectedGameChart(filteredPlayers, game) {
  const data = aggregateSum(filteredPlayers, 'country', 'winnings').slice(0, 8);
  const context = document.getElementById('selectedGameChart');

  createOrUpdateChart('selectedGameChart', context, {
    type: 'bar',
    data: {
      labels: data.map((d) => d.label),
      datasets: [
        {
          data: data.map((d) => d.value),
          backgroundColor: '#8b5cf6',
          borderRadius: 6
        }
      ]
    },
    options: {
      ...baseChartOptions(`Top countries for ${game}`),
      scales: {
        x: {
          ticks: {
            callback(value) {
              return compactMoney(value);
            },
            color: '#4a4a4a'
          },
          grid: {
            color: '#e5e5df'
          }
        },
        y: {
          ticks: {
            color: '#4a4a4a'
          },
          grid: {
            display: false
          }
        }
      },
      plugins: {
        title: {
          display: true,
          text: `Top countries for ${game}`,
          color: '#1f1f1f',
          font: { size: 16, weight: '600' },
          padding: { bottom: 16 }
        },
        legend: {
          display: false
        },
        tooltip: {
          callbacks: {
            label(context) {
              return `Total winnings: ${money(context.raw)}`;
            }
          }
        }
      }
    }
  });
}

function renderTournamentChart(filteredTournaments, game) {
  const context = document.getElementById('tournamentChart');
  const plotted = filteredTournaments
    .filter((row) => row.players > 0 && row.prizeMoney > 0)
    .sort((a, b) => b.prizeMoney - a.prizeMoney)
    .slice(0, 25)
    .map((row) => ({
      x: row.players,
      y: row.prizeMoney,
      r: Math.max(5, Math.min(14, (row.teams || 4) / 2)),
      tournamentName: row.tournamentName,
      teams: row.teams
    }));

  const averagePrize = filteredTournaments.length
    ? filteredTournaments.reduce((sum, row) => sum + row.prizeMoney, 0) / filteredTournaments.length
    : 0;

  document.getElementById('tournament-note').textContent = filteredTournaments.length
    ? `${game} has ${filteredTournaments.length} tournaments in this dataset, with an average listed prize pool of ${money(averagePrize)}.`
    : `No tournament data is currently available for ${game}.`;

  createOrUpdateChart('tournamentChart', context, {
    type: 'bubble',
    data: {
      datasets: [
        {
          label: game,
          data: plotted,
          backgroundColor: 'rgba(59, 130, 246, 0.45)',
          borderColor: '#2563eb',
          borderWidth: 1
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: {
          display: false
        },
        title: {
          display: true,
          text: `Tournament prize money vs. players for ${game}`,
          color: '#1f1f1f',
          font: { size: 16, weight: '600' },
          padding: { bottom: 16 }
        },
        tooltip: {
          callbacks: {
            title(context) {
              return context[0].raw.tournamentName;
            },
            label(context) {
              return [
                `Players: ${context.raw.x}`,
                `Prize money: ${money(context.raw.y)}`,
                `Teams: ${context.raw.teams || 'N/A'}`
              ];
            }
          }
        }
      },
      scales: {
        x: {
          title: {
            display: true,
            text: 'Number of players',
            color: '#4a4a4a'
          },
          ticks: {
            color: '#4a4a4a'
          },
          grid: {
            color: '#e5e5df'
          }
        },
        y: {
          title: {
            display: true,
            text: 'Prize money (USD)',
            color: '#4a4a4a'
          },
          ticks: {
            callback(value) {
              return compactMoney(value);
            },
            color: '#4a4a4a'
          },
          grid: {
            color: '#e5e5df'
          }
        }
      }
    }
  });
}

function renderConcentration() {
  const data = aggregateSum(state.players, 'game', 'winnings');
  const total = data.reduce((sum, row) => sum + row.value, 0);
  const topFive = data.slice(0, 5);

  document.getElementById('concentrationList').innerHTML = topFive
    .map((row) => {
      const share = total ? row.value / total : 0;
      return `
        <div class="concentration-item">
          <div class="concentration-topline">
            <span>${escapeHtml(row.label)}</span>
            <span>${percent(share)} of total player winnings</span>
          </div>
          <div class="concentration-track">
            <div class="concentration-fill" style="width: ${Math.max(share * 100, 2)}%"></div>
          </div>
        </div>
      `;
    })
    .join('');

  const topThreeShare = topFive.slice(0, 3).reduce((sum, row) => sum + row.value, 0) / total;
  document.getElementById('final-takeaway').textContent = `Across the full dataset, the top three games account for ${percent(topThreeShare)} of total player winnings. Even though esports appears broad, the financial structure shown here is much more concentrated than it first seems.`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll('&', '&amp;')
    .replaceAll('<', '&lt;')
    .replaceAll('>', '&gt;')
    .replaceAll('"', '&quot;')
    .replaceAll("'", '&#39;');
}
