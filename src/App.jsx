// AI response block quote style
const airesponse_bq = {
  fontSize: 14,
  lineHeight: '20px',
  paddingTop: 13,
  paddingBottom: 15,
  paddingLeft: 24,
  marginTop: 4,
  marginBottom: 12,
  color: '#424242',
  fontStyle: 'normal',
  borderLeft: '1px solid #707070',
};
// Grouped AI response list styles
const AIResponseListStyles = {
  ul: {
    fontSize: 14,
    lineHeight: '20px',
    color: '#424242',
    fontWeight: 400,
    marginTop: 0,
    marginBottom: 0,
    paddingLeft: 20,
    paddingTop: 2,
    paddingBottom: 0,
    listStyleType: 'disc',
    marginLeft: 0,
  },
  ol: {
    fontSize: 14,
    lineHeight: '20px',
    color: '#424242',
    fontWeight: 400,
    marginTop: 0,
    marginBottom: 0,
    paddingLeft: 32, // increased for marker visibility
    paddingTop: 2,
    paddingBottom: 4,
    listStyleType: 'decimal',
    listStylePosition: 'outside',
  },
  nested_ol: {
    fontSize: 14,
    lineHeight: '20px',
    color: '#424242',
    fontWeight: 400,
    marginTop: 0,
    marginBottom: 0,
    paddingLeft: 28, // slightly less than top-level ol
    listStyleType: 'lower-roman',
    listStylePosition: 'outside',
  },
};
// AI response paragraph style
const airesponse_p = {
  fontSize: 14,
  lineHeight: '20px',
  color: '#424242',
  fontWeight: 400,
  paddingTop: 2,
  paddingBottom: 4,
};
// Grouped AI response heading styles
const aiResponseHeadingStyles = {
  h1_large: {
    fontSize: 24,
    lineHeight: '32px',
    fontWeight: 600, // display semibold
    color: '#424242',
    fontFamily: 'Segoe UI Variable Display Semibold, Segoe UI, Segoe Sans, Arial, sans-serif',
    paddingTop: 6,
    paddingBottom: 0,
    marginTop: 0,
    marginBottom: 0,
  },
  h2_medium: {
    fontSize: 20,
    lineHeight: '28px',
    fontWeight: 600, // display semibold
    color: '#424242',
    fontFamily: 'Segoe UI Variable Display Semibold, Segoe UI, Segoe Sans, Arial, sans-serif',
    paddingTop: 6,
    paddingBottom: 0,
    marginTop: 0,
    marginBottom: 0,
  },
  h3_small: {
    fontSize: 18,
    lineHeight: '24px',
    fontWeight: 600,
    color: '#424242',
    fontFamily: 'Segoe UI Variable Display Semibold, Segoe UI, Segoe Sans, Arial, sans-serif',
    paddingTop: 6,
    paddingBottom: 0,
    marginTop: 0,
    marginBottom: 0,
  },
  h4_extrasmall: {
    fontSize: 16,
    lineHeight: '24px',
    fontWeight: 600,
    color: '#424242',
    paddingTop: 4,
    paddingBottom: 0,
  },
  h5_tiny: {
    fontSize: 14,
    lineHeight: '20px',
    fontWeight: 600,
    color: '#424242',
    marginTop: 0,
    marginBottom: 0,
    paddingTop: 4,
    paddingBottom: 0,
  },
  h6_micro: {
    fontSize: 14,
    lineHeight: '20px',
    fontWeight: 400,
    color: '#424242',
    marginTop: 0,
    marginBottom: 0,
    paddingTop: 4,
    paddingBottom: 0,
  },
};
import * as React from 'react';
import {
  Button,
  DrawerBody,
  DrawerHeader,
  Drawer,
  Input,
  makeStyles,
  shorthands,
  tokens,
} from '@fluentui/react-components';
import { Dismiss24Regular, Add24Regular, Eye24Regular, Search24Regular, CalendarLtr24Regular, MoreHorizontal24Regular, TextAlignLeft24Regular, ShieldTask24Regular, ChevronDown24Regular, Compose24Regular, Briefcase24Filled, Globe24Regular } from '@fluentui/react-icons';


// Type settings for Copilot pane (edit here to tinker)
const COPILOT_TYPE = {
  fontFamily: 'Segoe UI',
  fontSize: '1.5rem',
  fontWeight: 700,
  lineHeight: '1.2',
  padding: '0',
};

const RIBBON_HEIGHT = 48;
const useStyles = makeStyles({
  appRoot: {
    fontFamily: 'Segoe UI, Arial, sans-serif',
    background: '#ECECEC',
    minHeight: '100vh',
    ...shorthands.padding('0'),
  },
  ribbon: {
    position: 'fixed',
    top: 0,
    left: 0,
    width: '100vw',
    background: '#F8F8F8',
    color: '#222',
    boxSizing: 'border-box',
    zIndex: 200,
    borderBottom: '1.5px solid #E1E1E1',
    boxShadow: '0 1px 0 #E1E1E1',
    fontFamily: 'Segoe UI, Arial, sans-serif',
  },
  ribbonTitle: {
    fontWeight: 700,
    fontSize: '1.2rem',
    marginRight: '32px',
  },
  ribbonSpacer: {
    flex: 1,
  },
  docArea: {
    margin: '0 auto',
    maxWidth: '900px',
    minWidth: '680px',
    minHeight: '900px',
    background: '#ECECEC',
    borderRadius: 8,
    boxShadow: '0 4px 24px 0 rgba(0,0,0,0.13)',
    padding: '40px 48px',
    position: 'relative',
    marginTop: `${RIBBON_HEIGHT + 32}px`,
    display: 'flex',
    flexDirection: 'column',
    alignItems: 'center',
    zIndex: 1,
  marginRight: 0,
    border: '1.5px solid #E1E1E1',
  },
  copilotButton: {
    marginLeft: 'auto',
  },
});

function App() {
  const [copilotOpen, setCopilotOpen] = React.useState(true); // Open by default for prototyping
  const [copilotState, setCopilotState] = React.useState('default'); // 'default' or 'submitted'
  const [userMessage, setUserMessage] = React.useState('');
  const [messageCount, setMessageCount] = React.useState(0);
  const styles = useStyles();

  // Filler text for randomized responses
  const filler = {
    h1: 'Satya Nadella: A Visionary Leader',
    h2: 'Microsoft Achievements',
    h3: 'Cloud Transformation',
    h4: 'Executive Summary',
    h5: 'Early Life',
    h6: 'Additional Info',
    p: [
      'Satya Nadella has transformed Microsoft into a cloud-first company. His leadership style emphasizes empathy and innovation. Under his guidance, Microsoft has expanded its global reach and impact.',
      'Microsoft has achieved significant milestones in artificial intelligence. The company continues to invest in cutting-edge technologies. Nadella encourages a culture of continuous learning and growth.',
      'Nadella believes in empowering every person and organization. He has fostered a collaborative environment at Microsoft. The company’s success is attributed to its diverse and talented workforce.'
    ],
    ul: [
      'Creaming: Beating butter and sugar together until light and fluffy. This incorporates air and helps with leavening.',
      'Folding: Gently combining ingredients to retain air and volume in the mixture.',
      'Proofing: Allowing dough to rest and rise, developing flavor and texture.'
    ],
    ol: ['Joined Microsoft in 1992', 'Became CEO in 2014', 'Became Chairman in 2021'],
    bq: '“Success is not just about innovation, but about transformation.”',
    nested_ol: [
      {
        text: 'Test conditions',
        nested: [
          'Real-world constraints'
        ]
      },
      {
        text: 'Expected feedback',
        nested: []
      }
    ],
  };

  // Sibling selector pairings (headings never last)
  const headingKeys = ['h1', 'h2', 'h3', 'h4', 'h5', 'h6'];
  const pairings = [
    ['h1', 'p'],
    ['h3', 'p'],
    ['h3', 'h5', 'p'],
    ['h4', 'h5', 'p'],
    ['p', 'p'],
    ['p', 'ul'],
    ['p', 'ol'],
    ['ul', 'p'],
    ['ol', 'p'],
    ['p', 'bq'],
    ['h3', 'ul'],
    ['h4', 'ol'],
    ['ul', 'ul'],
    ['ol', 'ol'],
    ['nested_ol'],
  ];

  // Helper to get style by key
  const getStyle = (key) => {
    if (key === 'h1') return aiResponseHeadingStyles.h1_large;
    if (key === 'h2') return aiResponseHeadingStyles.h2_medium;
    if (key === 'h3') return aiResponseHeadingStyles.h3_small;
    if (key === 'h4') return aiResponseHeadingStyles.h4_extrasmall;
    if (key === 'h5') return aiResponseHeadingStyles.h5_tiny;
    if (key === 'h6') return aiResponseHeadingStyles.h6_micro;
    if (key === 'p') return airesponse_p;
    if (key === 'ul') return AIResponseListStyles.ul;
    if (key === 'ol') return AIResponseListStyles.ol;
    if (key === 'bq') return airesponse_bq;
    return {};
  };

  // Helper to render element by key
  const renderElement = (key, idx) => {
    if (key === 'ul') {
      return (
        <ul style={getStyle('ul')} key={idx}>
          {filler.ul.map((item, i) => {
            const firstColon = item.indexOf(':');
            let firstWord = '', rest = '';
            if (firstColon > 0) {
              firstWord = item.slice(0, firstColon + 1);
              rest = item.slice(firstColon + 1);
            } else {
              const firstSpace = item.indexOf(' ');
              firstWord = item.slice(0, firstSpace > 0 ? firstSpace : undefined);
              rest = item.slice(firstWord.length);
            }
            return (
              <li key={i} style={{ paddingBottom: 4, marginLeft: 8 }}>
                <span style={{ fontWeight: 700 }}>{firstWord}</span>
                <span>{rest}</span>
              </li>
            );
          })}
        </ul>
      );
    }
    if (key === 'ol') {
      return (
        <ol style={getStyle('ol')} key={idx}>
          {filler.ol.map((item, i) => (
            <li key={i} style={{ paddingBottom: 0 }}>
              <span style={{ marginLeft: 0 }}>{item}</span>
            </li>
          ))}
        </ol>
      );
    }
    if (key === 'bq') {
      // Remove any extra quotes from the filler text
      const text = filler.bq.replace(/^['"“”‘’]+|['"“”‘’]+$/g, '');
      return <div style={getStyle('bq')} key={idx}>&ldquo;{text}&rdquo;</div>;
    }
    if (key === 'nested_ol') {
      return (
        <ol style={getStyle('ol')} key={idx}>
          {filler.nested_ol.map((item, i) => (
            <li key={i} style={{ paddingBottom: 0 }}>
              <span style={{ marginLeft: 0 }}>{item.text}</span>
              {item.nested && item.nested.length > 0 && (
                <ol type="i" style={{ ...getStyle('nested_ol'), marginTop: 0, marginBottom: 0 }}>
                  {item.nested.map((sub, j) => (
                    <li key={j} style={{ paddingBottom: 0 }}>
                      <span style={{ marginLeft: 0 }}>{sub}</span>
                    </li>
                  ))}
                </ol>
              )}
            </li>
          ))}
        </ol>
      );
    }
    if (key === 'p') {
      // Pick a random paragraph from the array
      const sentences = filler.p[Math.floor(Math.random() * filler.p.length)];
      return <div style={getStyle('p')} key={idx}>{sentences}</div>;
    }
    return <div style={getStyle(key)} key={idx}>{filler[key]}</div>;
  };

  // Number of elements in the static response (biography)
  const staticResponseLength = 11;

  // Generate a random response using multiple sibling selector pairings
  const getRandomResponse = () => {
    let total = 0;
    let keys = [];
    while (total < staticResponseLength - 1) { // leave space for bq
      const pairing = pairings[Math.floor(Math.random() * pairings.length)];
      keys = keys.concat(pairing);
      total = keys.length;
    }
    // If we have more than needed, trim to exact length minus 1
    keys = keys.slice(0, staticResponseLength - 1);
    // Ensure last element before bq is not a heading
    if (headingKeys.includes(keys[keys.length - 1])) {
      for (let i = keys.length - 2; i >= 0; i--) {
        if (!headingKeys.includes(keys[i])) {
          [keys[i], keys[keys.length - 1]] = [keys[keys.length - 1], keys[i]];
          break;
        }
      }
    }
    // Always add block quote at the end
    keys.push('bq');
    return keys.map((key, idx) => renderElement(key, idx));
  };

  // Handler for sending message
  const handleSend = (msg) => {
    setUserMessage(msg);
    setCopilotState('submitted');
    setMessageCount((c) => c + 1);
  };

  return (
    <div className={styles.appRoot}>
      {/* Ribbon */}
      <div className={styles.ribbon}>
        {/* Simulated Word Web Ribbon */}
        <div style={{
          display: 'flex',
          alignItems: 'center',
          height: 44,
          borderBottom: '1.5px solid #E1E1E1',
          padding: '0 24px',
          background: '#F8F8F8',
        }}>
          {/* App/Word icon */}
          <div style={{ width: 32, height: 32, background: '#2564cf', borderRadius: 6, display: 'flex', alignItems: 'center', justifyContent: 'center', marginRight: 16 }}>
            <span style={{ color: '#fff', fontWeight: 700, fontSize: 18, fontFamily: 'Segoe UI' }}>W</span>
          </div>
          {/* Tabs */}
          <div style={{ display: 'flex', alignItems: 'center', gap: 0 }}>
            <span style={{ fontWeight: 600, fontSize: 15, marginRight: 32, color: '#222', borderBottom: '2.5px solid #2564cf', paddingBottom: 2 }}>Home</span>
            <span style={{ marginRight: 24, color: '#222', opacity: 0.85 }}>Insert</span>
            <span style={{ marginRight: 24, color: '#222', opacity: 0.85 }}>Draw</span>
            <span style={{ marginRight: 24, color: '#222', opacity: 0.85 }}>Design</span>
            <span style={{ marginRight: 24, color: '#222', opacity: 0.85 }}>Layout</span>
            <span style={{ marginRight: 24, color: '#222', opacity: 0.85 }}>References</span>
            <span style={{ marginRight: 24, color: '#222', opacity: 0.85 }}>Mailings</span>
            <span style={{ marginRight: 24, color: '#222', opacity: 0.85 }}>Review</span>
            <span style={{ marginRight: 24, color: '#222', opacity: 0.85 }}>View</span>
          </div>
          <div style={{ flex: 1 }} />
          {/* Copilot Pane Toggle Button - neutral style */}
          <Button
            appearance="subtle"
            style={{ marginLeft: 16, background: 'transparent', color: '#222', fontWeight: 500, fontSize: 15, display: 'flex', alignItems: 'center', gap: 8 }}
            onClick={() => setCopilotOpen((o) => !o)}
          >
            Copilot Pane
            <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" style={{ width: 20, height: 20 }}>
              <path d="M17.1941 3.72897C16.8913 2.84433 16.0596 2.25 15.1245 2.25L14.0807 2.25C13.0219 2.25 12.1151 3.0083 11.9277 4.05038L10.8384 10.1086L11.2624 8.6872C11.5389 7.76028 12.3913 7.125 13.3586 7.125L17.2947 7.125L18.9792 8.10676L20.6031 7.125H19.92C18.985 7.125 18.1533 6.53067 17.8504 5.64603L17.1941 3.72897Z" fill="url(#paint0_radial_890_1287)"/>
              <path d="M7.04739 20.2611C7.34721 21.1509 8.18142 21.7501 9.12037 21.7501H10.9296C12.133 21.7501 13.1105 20.7781 13.1171 19.5747L13.1487 13.8647L12.7275 15.3026C12.4545 16.2347 11.5995 16.8751 10.6282 16.8751H6.6673L5.22338 15.7661L3.66016 16.8751H4.33521C5.27415 16.8751 6.10837 17.4743 6.40819 18.3641L7.04739 20.2611Z" fill="url(#paint1_radial_890_1287)"/>
              <path d="M14.8127 2.25H6.60958C4.26585 2.25 2.8596 5.31687 1.92211 8.38374C0.811419 12.0172 -0.641941 16.8766 3.56272 16.8766H7.24547C8.22122 16.8766 9.07882 16.2311 9.3494 15.2936C9.95889 13.1819 11.0626 9.37449 11.9262 6.48894C12.357 5.04937 12.7158 3.81303 13.2665 3.04312C13.5753 2.61147 14.0899 2.25 14.8127 2.25Z" fill="url(#paint2_radial_890_1287)"/>
              <path d="M14.8127 2.25H6.60958C4.26585 2.25 2.8596 5.31687 1.92211 8.38374C0.811419 12.0172 -0.641941 16.8766 3.56272 16.8766H7.24547C8.22122 16.8766 9.07882 16.2311 9.3494 15.2936C9.95889 13.1819 11.0626 9.37449 11.9262 6.48894C12.357 5.04937 12.7158 3.81303 13.2665 3.04312C13.5753 2.61147 14.0899 2.25 14.8127 2.25Z" fill="url(#paint3_linear_890_1287)"/>
              <path d="M9.18604 21.75H17.3891C19.7329 21.75 21.1391 18.6827 22.0766 15.6155C23.1873 11.9816 24.6406 7.12158 20.436 7.12158H16.7533C15.7775 7.12158 14.9199 7.76711 14.6493 8.70462C14.0399 10.8166 12.9361 14.6246 12.0725 17.5105C11.6417 18.9503 11.2829 20.1868 10.7322 20.9568C10.4234 21.3885 9.90882 21.75 9.18604 21.75Z" fill="url(#paint4_radial_890_1287)"/>
              <path d="M9.18604 21.75H17.3891C19.7329 21.75 21.1391 18.6827 22.0766 15.6155C23.1873 11.9816 24.6406 7.12158 20.436 7.12158H16.7533C15.7775 7.12158 14.9199 7.76711 14.6493 8.70462C14.0399 10.8166 12.9361 14.6246 12.0725 17.5105C11.6417 18.9503 11.2829 20.1868 10.7322 20.9568C10.4234 21.3885 9.90882 21.75 9.18604 21.75Z" fill="url(#paint5_linear_890_1287)"/>
              <defs>
                <radialGradient id="paint0_radial_890_1287" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="translate(19.1812 10.16) rotate(-130.79) scale(8.47051 8.0375)"><stop offset="0.0955758" stopColor="#00AEFF"/><stop offset="0.773185" stopColor="#2253CE"/><stop offset="1" stopColor="#0736C4"/></radialGradient>
                <radialGradient id="paint1_radial_890_1287" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="translate(5.38158 16.4612) rotate(50.1613) scale(7.74454 7.6066)"><stop stopColor="#FFB657"/><stop offset="0.633728" stopColor="#FF5F3D"/><stop offset="0.923392" stopColor="#C02B3C"/></radialGradient>
                <radialGradient id="paint2_radial_890_1287" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="translate(6.4063 16.8714) rotate(-93.282) scale(12.9812 73.6372)"><stop offset="0.03" stopColor="#FFC800"/><stop offset="0.31" stopColor="#98BD42"/><stop offset="0.49" stopColor="#52B471"/><stop offset="0.843838" stopColor="#0D91E1"/></radialGradient>
                <linearGradient id="paint3_linear_890_1287" x1="7.14149" y1="2.25" x2="7.76802" y2="16.8771" gradientUnits="userSpaceOnUse"><stop stopColor="#3DCBFF"/><stop offset="0.246674" stopColor="#0588F7" stopOpacity="0"/></linearGradient>
                <radialGradient id="paint4_radial_890_1287" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="translate(20.8574 5.68936) rotate(109.453) scale(19.4586 23.4934)"><stop offset="0.0661714" stopColor="#8C48FF"/><stop offset="0.5" stopColor="#F2598A"/><stop offset="0.895833" stopColor="#FFB152"/></radialGradient>
                <linearGradient id="paint5_linear_890_1287" x1="21.5054" y1="6.22849" x2="21.4972" y2="10.2127" gradientUnits="userSpaceOnUse"><stop offset="0.0581535" stopColor="#F8ADFA"/><stop offset="0.708063" stopColor="#A86EDD" stopOpacity="0"/></linearGradient>
              </defs>
            </svg>
          </Button>
        </div>
        {/* Ribbon command row (simulated, not interactive) */}
        <div style={{
          display: 'flex',
          alignItems: 'center',
          height: 40,
          position: 'relative',
          width: '100vw',
          background: '#F3F3F3',
          padding: '0 24px',
          borderBottom: '1.5px solid #E1E1E1',
          boxSizing: 'border-box',
          overflowX: 'auto',
        }}>
          {/* Simulated command buttons */}
          <Button appearance="subtle" style={{ marginRight: 8, fontSize: 15, color: '#222', background: 'transparent' }} icon={<Add24Regular />} />
          <Button appearance="subtle" style={{ marginRight: 8, fontSize: 15, color: '#222', background: 'transparent' }} icon={<Eye24Regular />} />
          <Button appearance="subtle" style={{ marginRight: 8, fontSize: 15, color: '#222', background: 'transparent' }} icon={<CalendarLtr24Regular />} />
          <Button appearance="subtle" style={{ marginRight: 8, fontSize: 15, color: '#222', background: 'transparent' }} icon={<Search24Regular />} />
          <Button appearance="subtle" style={{ marginRight: 8, fontSize: 15, color: '#222', background: 'transparent' }} icon={<MoreHorizontal24Regular />} />
          <div style={{ flex: 1 }} />
          <Button appearance="subtle" style={{ marginRight: 8, fontSize: 15, color: '#222', background: 'transparent' }}>Comments</Button>
          <Button appearance="subtle" style={{ marginRight: 8, fontSize: 15, color: '#222', background: 'transparent' }}>Editing</Button>
          <Button appearance="primary" style={{ background: '#2564cf', color: '#fff', fontSize: 15, borderRadius: 6, minWidth: 60 }}>Share</Button>
        </div>
      </div>

      {/* Document Area */}
      <div className={styles.docArea}>
        {/* Simulate Word's document area */}
        <div style={{ width: '100%', height: 40, background: '#f3f3f3', borderRadius: 6, marginBottom: 16, display: 'flex', alignItems: 'center', justifyContent: 'center', color: '#888', fontSize: 15, fontWeight: 500 }}>
          <span>What do you want Copilot to draft?</span>
        </div>
        <div style={{ background: '#fff', minHeight: 700, border: '1.5px solid #e0e0e0', borderRadius: 8, width: '100%' }} />
      </div>

      {/* Copilot Side Pane */}
      <Drawer
        open={copilotOpen}
        position="end"
        onOpenChange={(_, d) => setCopilotOpen(d.open)}
        style={{
          boxShadow: '0 0 16px 0 rgba(0,0,0,0.08)',
          borderLeft: '1.5px solid #e0e0e0',
          background: '#FAFAFA',
          width: 350,
          minWidth: 350,
          maxWidth: 350,
          paddingRight: 0,
          top: '84px',
          height: 'calc(100vh - 84px)',
          position: 'fixed',
          right: 0,
          display: 'flex',
          flexDirection: 'column',
          zIndex: 300,
        }}
      >
        {/* Copilot Header */}
        <div style={{
          background: '#ECECEC',
          padding: '4px 20px 4px 20px',
          display: 'flex',
          alignItems: 'center',
          justifyContent: 'space-between',
          minHeight: 28,
          height: 36,
          width: '100%',
          boxSizing: 'border-box',
        }}>
          <div style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" style={{ width: 18, height: 18 }}>
              <path d="M17.1941 3.72897C16.8913 2.84433 16.0596 2.25 15.1245 2.25L14.0807 2.25C13.0219 2.25 12.1151 3.0083 11.9277 4.05038L10.8384 10.1086L11.2624 8.6872C11.5389 7.76028 12.3913 7.125 13.3586 7.125L17.2947 7.125L18.9792 8.10676L20.6031 7.125H19.92C18.985 7.125 18.1533 6.53067 17.8504 5.64603L17.1941 3.72897Z" fill="url(#paint0_radial_890_1287)"/>
              <path d="M7.04739 20.2611C7.34721 21.1509 8.18142 21.7501 9.12037 21.7501H10.9296C12.133 21.7501 13.1105 20.7781 13.1171 19.5747L13.1487 13.8647L12.7275 15.3026C12.4545 16.2347 11.5995 16.8751 10.6282 16.8751H6.6673L5.22338 15.7661L3.66016 16.8751H4.33521C5.27415 16.8751 6.10837 17.4743 6.40819 18.3641L7.04739 20.2611Z" fill="url(#paint1_radial_890_1287)"/>
              <path d="M14.8127 2.25H6.60958C4.26585 2.25 2.8596 5.31687 1.92211 8.38374C0.811419 12.0172 -0.641941 16.8766 3.56272 16.8766H7.24547C8.22122 16.8766 9.07882 16.2311 9.3494 15.2936C9.95889 13.1819 11.0626 9.37449 11.9262 6.48894C12.357 5.04937 12.7158 3.81303 13.2665 3.04312C13.5753 2.61147 14.0899 2.25 14.8127 2.25Z" fill="url(#paint2_radial_890_1287)"/>
              <path d="M14.8127 2.25H6.60958C4.26585 2.25 2.8596 5.31687 1.92211 8.38374C0.811419 12.0172 -0.641941 16.8766 3.56272 16.8766H7.24547C8.22122 16.8766 9.07882 16.2311 9.3494 15.2936C9.95889 13.1819 11.0626 9.37449 11.9262 6.48894C12.357 5.04937 12.7158 3.81303 13.2665 3.04312C13.5753 2.61147 14.0899 2.25 14.8127 2.25Z" fill="url(#paint3_linear_890_1287)"/>
              <path d="M9.18604 21.75H17.3891C19.7329 21.75 21.1391 18.6827 22.0766 15.6155C23.1873 11.9816 24.6406 7.12158 20.436 7.12158H16.7533C15.7775 7.12158 14.9199 7.76711 14.6493 8.70462C14.0399 10.8166 12.9361 14.6246 12.0725 17.5105C11.6417 18.9503 11.2829 20.1868 10.7322 20.9568C10.4234 21.3885 9.90882 21.75 9.18604 21.75Z" fill="url(#paint4_radial_890_1287)"/>
              <path d="M9.18604 21.75H17.3891C19.7329 21.75 21.1391 18.6827 22.0766 15.6155C23.1873 11.9816 24.6406 7.12158 20.436 7.12158H16.7533C15.7775 7.12158 14.9199 7.76711 14.6493 8.70462C14.0399 10.8166 12.9361 14.6246 12.0725 17.5105C11.6417 18.9503 11.2829 20.1868 10.7322 20.9568C10.4234 21.3885 9.90882 21.75 9.18604 21.75Z" fill="url(#paint5_linear_890_1287)"/>
              <defs>
                <radialGradient id="paint0_radial_890_1287" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="translate(19.1812 10.16) rotate(-130.79) scale(8.47051 8.0375)"><stop offset="0.0955758" stopColor="#00AEFF"/><stop offset="0.773185" stopColor="#2253CE"/><stop offset="1" stopColor="#0736C4"/></radialGradient>
                <radialGradient id="paint1_radial_890_1287" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="translate(5.38158 16.4612) rotate(50.1613) scale(7.74454 7.6066)"><stop stopColor="#FFB657"/><stop offset="0.633728" stopColor="#FF5F3D"/><stop offset="0.923392" stopColor="#C02B3C"/></radialGradient>
                <radialGradient id="paint2_radial_890_1287" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="translate(6.4063 16.8714) rotate(-93.282) scale(12.9812 73.6372)"><stop offset="0.03" stopColor="#FFC800"/><stop offset="0.31" stopColor="#98BD42"/><stop offset="0.49" stopColor="#52B471"/><stop offset="0.843838" stopColor="#0D91E1"/></radialGradient>
                <linearGradient id="paint3_linear_890_1287" x1="7.14149" y1="2.25" x2="7.76802" y2="16.8771" gradientUnits="userSpaceOnUse"><stop stopColor="#3DCBFF"/><stop offset="0.246674" stopColor="#0588F7" stopOpacity="0"/></linearGradient>
                <radialGradient id="paint4_radial_890_1287" cx="0" cy="0" r="1" gradientUnits="userSpaceOnUse" gradientTransform="translate(20.8574 5.68936) rotate(109.453) scale(19.4586 23.4934)"><stop offset="0.0661714" stopColor="#8C48FF"/><stop offset="0.5" stopColor="#F2598A"/><stop offset="0.895833" stopColor="#FFB152"/></radialGradient>
                <linearGradient id="paint5_linear_890_1287" x1="21.5054" y1="6.22849" x2="21.4972" y2="10.2127" gradientUnits="userSpaceOnUse"><stop offset="0.0581535" stopColor="#F8ADFA"/><stop offset="0.708063" stopColor="#A86EDD" stopOpacity="0"/></linearGradient>
              </defs>
            </svg>
            <span style={{ fontWeight: 700, fontSize: 14, lineHeight: '20px', letterSpacing: '-0.2px', color: '#222' }}>Copilot</span>
          </div>
          <Button
            appearance="subtle"
            icon={<Dismiss24Regular style={{ fontSize: 14 }} />}
            onClick={() => setCopilotOpen(false)}
            aria-label="Close Copilot"
            style={{ color: '#666', borderRadius: '50%', minWidth: 22, minHeight: 22, padding: 0 }}
          />
        </div>
        {/* Copilot Toolbar */}
  <div style={{ width: '100%', background: '#FAFAFA', padding: '16px 24px', boxSizing: 'border-box' }}>
          <div style={{
            display: 'grid',
            gridTemplateColumns: 'min-content 1fr min-content',
            alignItems: 'center',
            width: '100%',
            gap: 8
          }}>
            {/* Left group */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 0, minWidth: 96 }}>
              <Button appearance="subtle" style={{ color: '#222', minWidth: 32, minHeight: 32, width: 32, height: 32, borderRadius: 6 }} icon={<TextAlignLeft24Regular style={{ fontSize: 18 }} />} aria-label="Text Align" />
              <Button appearance="subtle" style={{ color: '#1a7f37', minWidth: 32, minHeight: 32, width: 32, height: 32, borderRadius: 6 }} icon={<ShieldTask24Regular style={{ fontSize: 18 }} />} aria-label="Secure" />
              {/* Spacer to match right group width (2 buttons + gap) */}
              <div style={{ width: 32, height: 32, opacity: 0, pointerEvents: 'none' }} />
            </div>
            {/* Center toggle */}
            <div style={{ display: 'flex', alignItems: 'center', justifyContent: 'center' }}>
              <div style={{ display: 'flex', alignItems: 'center', border: '2px solid #e0e0e0', borderRadius: 10, overflow: 'hidden', height: 32 }}>
                <Button appearance="subtle" style={{ background: '#f5f8fd', color: '#0F6CBD', minWidth: 32, minHeight: 32, width: 32, height: 32, borderRadius: 0, fontWeight: 600, fontSize: 16 }} icon={<Briefcase24Filled style={{ fontSize: 20, color: '#0F6CBD' }} />} aria-label="Briefcase" />
                <Button appearance="subtle" style={{ background: '#fff', color: '#222', minWidth: 32, minHeight: 32, width: 32, height: 32, borderRadius: 0, fontWeight: 600, fontSize: 16 }} icon={<Globe24Regular style={{ fontSize: 20, color: '#222' }} />} aria-label="Globe" />
              </div>
            </div>
            {/* Right group */}
            <div style={{ display: 'flex', alignItems: 'center', gap: 8, minWidth: 96 }}>
              {/* Spacer to match left group width (2 buttons + gap) */}
              <div style={{ width: 32, height: 32, opacity: 0, pointerEvents: 'none' }} />
              <Button appearance="primary" style={{ background: '#0F6CBD', color: '#fff', minWidth: 32, minHeight: 32, width: 32, height: 32, borderRadius: 8 }} icon={<Compose24Regular style={{ fontSize: 20 }} />} aria-label="Compose" />
              <Button appearance="subtle" style={{ color: '#222', minWidth: 32, minHeight: 32, width: 32, height: 32, borderRadius: 6 }} icon={<MoreHorizontal24Regular />} aria-label="More" />
            </div>
          </div>
        </div>
        <DrawerBody style={{ padding: '0 24px 0 24px', display: 'flex', flexDirection: 'column', flex: 1, minHeight: 0 }}>
          {/* Copilot Main Message and Suggestions (default state) */}
          {copilotState === 'default' && <>
            <div style={{
              fontFamily: 'Segoe UI Variable Display Semibold, Segoe UI, Segoe Sans, Arial, sans-serif',
              fontSize: 24,
              fontWeight: 600,
              lineHeight: '32px',
              padding: COPILOT_TYPE.padding,
              marginBottom: 48,
              marginTop: 80,
              textAlign: 'center',
              letterSpacing: '-0.5px',
            }}>
              Let's organize your thoughts
            </div>
            <CopilotInputBox onSend={handleSend} />
            <div style={{ display: 'flex', flexDirection: 'column', gap: 16, marginBottom: 16 }}>
              <div style={{ background: '#fff', borderRadius: 12, boxShadow: '0 1px 4px #0001', padding: '10px 16px', border: '1.5px solid #e0e0e0', fontWeight: 500, fontSize: 12, color: '#222', cursor: 'pointer', minHeight: 32, display: 'flex', alignItems: 'center' }}>
                <span style={{ marginRight: 2 }}>Create an image of</span> <span style={{ color: '#1b4db1', fontWeight: 500, marginLeft: 2 }}>description</span>
              </div>
              <div style={{ background: '#fff', borderRadius: 12, boxShadow: '0 1px 4px #0001', padding: '10px 16px', border: '1.5px solid #e0e0e0', fontWeight: 500, fontSize: 12, color: '#222', cursor: 'pointer', minHeight: 32, display: 'flex', alignItems: 'center' }}>
                <span style={{ marginRight: 2 }}>List key points from</span> <span style={{ color: '#1b4db1', fontWeight: 500, marginLeft: 2 }}>Rocksteady Compete and Copilot benchmarking_Pt2...</span>
              </div>
              <div style={{ background: '#fff', borderRadius: 12, boxShadow: '0 1px 4px #0001', padding: '10px 16px', border: '1.5px solid #e0e0e0', fontWeight: 500, fontSize: 12, color: '#222', cursor: 'pointer', minHeight: 32, display: 'flex', alignItems: 'center' }}>
                <span style={{ marginRight: 2 }}>Help me brainstorm ideas for</span> <span style={{ color: '#1b4db1', fontWeight: 500, marginLeft: 2 }}>topic</span>
              </div>
            </div>
            <div style={{ color: '#888', fontSize: 12, lineHeight: '16px', marginTop: 0, marginBottom: 16, textAlign: 'right', cursor: 'pointer', display: 'flex', alignItems: 'center', justifyContent: 'flex-end' }}>
              See more <ChevronDown24Regular style={{ width: 12, height: 12, fontSize: 12, verticalAlign: 'middle', marginLeft: 4 }} />
            </div>
          </>}
          {/* Copilot submitted state: show user message under toolbar, input at bottom */}
          {copilotState === 'submitted' && (
            <>
              <div style={{ flex: 1, minHeight: 0, overflow: 'auto', display: 'flex', flexDirection: 'column', paddingRight: 24 }}>
                {/* ai conversation container */}
                <div style={{ width: '100%', display: 'flex', flexDirection: 'column', alignItems: 'flex-end', marginTop: 24 }}>
                  {/* Timestamp */}
                  <div style={{
                    textAlign: 'center',
                    fontSize: 12,
                    lineHeight: '16px',
                    color: '#707070',
                    marginTop: 4,
                    marginBottom: 12,
                    width: '100%',
                    fontWeight: 400,
                  }}>
                    Today
                  </div>
                  {/* User Prompt */}
                  <div style={{
                    background: '#EBEFFF',
                    borderRadius: '16px 16px 4px 16px',
                    padding: '8px 16px',
                    fontSize: 14,
                    lineHeight: '20px',
                    color: '#424242',
                    fontWeight: 400,
                    maxWidth: '80%',
                    wordBreak: 'break-word',
                    boxShadow: '0 1px 4px #0001',
                    alignSelf: 'flex-end',
                    marginTop: 20,
                    marginBottom: 16,
                    marginLeft: 32,
                    marginRight: 0,
                  }}>
                    {userMessage}
                  </div>
                  {/* Copilot Loader with Icon */}
                  <div style={{
                    display: 'flex',
                    alignItems: 'center',
                    alignSelf: 'flex-start',
                    marginLeft: 0,
                    marginTop: 0,
                    marginBottom: 4,
                  }}>
                    <svg xmlns="http://www.w3.org/2000/svg" width="24" height="24" viewBox="0 0 24 24" fill="none" style={{ display: 'block' }}>
                      <path d="M10.7789 18.9999H15.5938C16.6162 18.9999 17.3728 18.3436 17.9211 17.5283C18.4714 16.71 18.8793 15.638 19.1961 14.5875C19.5592 13.3835 20.0165 11.8667 19.9995 10.6448C19.991 10.0275 19.8622 9.40417 19.4636 8.93124C19.0503 8.44087 18.4256 8.2026 17.6254 8.2026H17.1916C16.8184 8.17148 16.4946 7.92276 16.3697 7.56318L15.936 6.31371C15.663 5.52708 14.9239 5 14.0939 5H13.2366L13.2211 5.00006H8.4062C7.38378 5.00006 6.62719 5.65641 6.07891 6.47169C5.52859 7.29001 5.12069 8.36198 4.80387 9.41253C4.44078 10.6165 3.98355 12.1333 4.00046 13.3552C4.009 13.9725 4.13784 14.5958 4.53642 15.0688C4.9497 15.5591 5.57443 15.7974 6.37461 15.7974H6.8084C7.18159 15.8285 7.50544 16.0772 7.63026 16.4368L8.06399 17.6863C8.33704 18.4729 9.07612 19 9.90608 19H10.7634C10.7686 19 10.7737 19 10.7789 18.9999ZM11.7753 10.6579L12.0257 9.87203C12.1521 9.4752 12.5196 9.20589 12.9347 9.20589H13.4642C13.3917 9.33608 13.3332 9.47609 13.2911 9.62418C13.0295 10.5428 12.63 11.9413 12.2246 13.3423L11.9743 14.128C11.8479 14.5248 11.4804 14.7941 11.0653 14.7941H10.5358C10.6083 14.6639 10.6668 14.5239 10.7089 14.3758C10.9705 13.4572 11.37 12.0589 11.7753 10.6579ZM11.5245 15.7392C11.3982 16.1671 11.2841 16.5504 11.1631 16.8883C11.0267 17.2694 10.8932 17.5582 10.7485 17.7632C10.6622 17.8855 10.4958 17.9999 10.1241 17.9999C10.1215 17.9999 10.1189 18 10.1163 18H9.90608C9.50003 18 9.13845 17.7421 9.00486 17.3573L8.57113 16.1078C8.53341 15.9992 8.4868 15.8954 8.43229 15.7974H8.8318C8.8699 15.7974 8.9078 15.7963 8.94546 15.7941L11.0653 15.7941C11.2226 15.7941 11.3765 15.7752 11.5245 15.7392ZM15.4289 7.89218C15.4666 8.00084 15.5132 8.10455 15.5677 8.2026H15.1682C15.1301 8.2026 15.0922 8.20371 15.0545 8.2059L12.9347 8.2059C12.7775 8.2059 12.6235 8.22481 12.4755 8.26079C12.6018 7.83287 12.7159 7.4496 12.8369 7.11166C12.9733 6.73061 13.1068 6.44177 13.2515 6.23683C13.3378 6.11449 13.5042 6.00005 13.8759 6.00005L13.8837 5.99999H14.0939C14.5 5.99999 14.8616 6.25786 14.9952 6.64271L15.4289 7.89218ZM5.75748 9.70223C6.06577 8.67999 6.43925 7.7233 6.90473 7.03115C7.37224 6.33597 7.86586 6.00005 8.4062 6.00005H12.2306C12.1043 6.23835 11.9967 6.5013 11.8993 6.77348C11.7618 7.15729 11.6343 7.58978 11.5027 8.03582L11.4736 8.1347C10.8956 10.0925 10.1567 12.676 9.75096 14.101C9.63358 14.5133 9.25841 14.7974 8.8318 14.7974H6.37461C5.77299 14.7974 5.46751 14.6251 5.29712 14.423C5.11203 14.2033 5.00376 13.8521 4.9967 13.3413C4.98239 12.3072 5.37998 10.954 5.75748 9.70223ZM17.0953 16.9689C16.6278 17.664 16.1341 17.9999 15.5938 17.9999H11.7694C11.8957 17.7617 12.0033 17.4987 12.1007 17.2265C12.2382 16.8427 12.3657 16.4102 12.4973 15.9642L12.5264 15.8653C13.1044 13.9075 13.8433 11.324 14.249 9.89895C14.3664 9.48671 14.7416 9.20259 15.1682 9.20259H17.6254C18.227 9.20259 18.5325 9.37487 18.7029 9.57705C18.888 9.79666 18.9962 10.1479 19.0033 10.6587C19.0176 11.6928 18.62 13.046 18.2425 14.2978C17.9342 15.32 17.5608 16.2767 17.0953 16.9689Z" fill="#242424"/>
                    </svg>
                    <span style={{
                      fontSize: 14,
                      lineHeight: '20px',
                      color: '#242424',
                      fontWeight: 600,
                      paddingTop: 2,
                      paddingBottom: 2,
                      marginLeft: 4,
                    }}>
                      Copilot
                    </span>
                  </div>
                  {/* AI response */}
                  <div style={{
                    width: '100%',
                    boxSizing: 'border-box',
                    paddingLeft: 0,
                    paddingRight: 0,
                    marginTop: 0,
                    marginLeft: 0,
                    marginRight: 0,
                    marginBottom: 0,
                    wordBreak: 'break-word',
                  }}>
                    {messageCount === 1 ? (
                      <>
                        <div style={airesponse_p}>
                          Sure, here is a short biography of Satya Nadella grouped by overview, details, and final thoughts:
                        </div>
                        {/* aiResponseHeadingStyles.h4_extrasmall */}
                        <div style={aiResponseHeadingStyles.h4_extrasmall}>Overview</div>
                        <div style={airesponse_p}>
                          Satya Nadella is the Chairman and CEO of Microsoft. He joined Microsoft in 1992 and has held various leadership roles, including executive vice president of Microsoft's cloud and enterprise group. Nadella became CEO in 2014 and chairman in 2021. He is recognized for leading Microsoft's transition to cloud computing and developing one of the largest cloud infrastructures in the world.
                        </div>
                        {/* aiResponseHeadingStyles.h4_extrasmall */}
                        <div style={aiResponseHeadingStyles.h4_extrasmall}>Details</div>
                        {/* aiResponseHeadingStyles.h5_tiny */}
                        <div style={aiResponseHeadingStyles.h5_tiny}>Early Life and Education</div>
                        <div style={airesponse_p}>
                          Satya Nadella was born on August 19, 1967, in Hyderabad, India. He moved to the United States to pursue higher education, earning a bachelor's degree in electrical engineering from the Manipal Institute of Technology, a master's degree in computer science from the University of Wisconsin–Milwaukee, and an MBA from the University of Chicago Booth School of Business.
                        </div>
                        {/* aiResponseHeadingStyles.h5_tiny */}
                        <div style={aiResponseHeadingStyles.h5_tiny}>Career at Microsoft</div>
                        <div style={airesponse_p}>
                          Nadella joined Microsoft in 1992 and has held various leadership roles within the company. He served as the executive vice president of Microsoft's cloud and enterprise group before becoming the CEO in 2014, succeeding Steve Ballmer. In 2021, he also became the chairman of Microsoft, succeeding John W. Thompson.
                        </div>
                        {/* aiResponseHeadingStyles.h4_extrasmall */}
                        <div style={aiResponseHeadingStyles.h4_extrasmall}>Final Thoughts</div>
                        <div style={airesponse_p}>
                          Satya Nadella's leadership has been instrumental in transforming Microsoft into a more innovative and forward-thinking company. His contributions to the company's transition to cloud computing and the development of a vast cloud infrastructure have been significant milestones in his career.
                        </div>
                        <div style={airesponse_p}>
                          Would you like to know more about his specific achievements or contributions to Microsoft?
                        </div>
                        {/* Example block quote */}
                        <div style={airesponse_bq}>
                          “The true test of leadership is how well you function in a crisis.”
                        </div>
                      </>
                    ) : (
                      getRandomResponse()
                    )}
                  </div>
                </div>
              </div>
              <div style={{ width: '100%', background: '#FAFAFA', paddingTop: 24, paddingBottom: 16 }}>
                <CopilotInputBox onSend={handleSend} resetOnSend />
              </div>
            </>
          )}
        </DrawerBody>
      </Drawer>
    </div>
  );
}

// Interactive AI input box component
import { Mic24Regular, ArrowRight24Regular } from '@fluentui/react-icons';

function CopilotInputBox({ onSend, resetOnSend }) {
  const [input, setInput] = React.useState("");
  // Reset input when resetOnSend is true and input is sent
  React.useEffect(() => {
    if (resetOnSend && input === "") setInput("");
  }, [resetOnSend, input]);

  const handleSend = () => {
    if (input.trim() && onSend) {
      onSend(input);
      if (resetOnSend) setInput("");
    }
  };

  return (
    <div style={{
      background: '#fff',
      borderRadius: 16,
      boxShadow: '0 2px 8px 0 rgba(0,0,0,0.08)',
      padding: '16px 16px 8px 16px',
      marginBottom: 24,
      border: '2px solid',
      borderImage: 'linear-gradient(90deg, #F5F5F5, #FAFAFA) 1',
      width: '100%',
      minHeight: 80,
      position: 'relative',
      display: 'flex',
      flexDirection: 'column',
      justifyContent: 'flex-start',
      boxSizing: 'border-box',
    }}>
      <input
        type="text"
        value={input}
        onChange={e => setInput(e.target.value)}
        placeholder="Message Copilot"
        style={{
          color: '#888',
          fontSize: 14,
          lineHeight: '20px',
          fontWeight: 500,
          border: 'none',
          outline: 'none',
          background: 'transparent',
          marginBottom: 24,
          width: '100%',
          padding: 0,
          boxShadow: 'none',
          borderRadius: 0,
          MozAppearance: 'none',
          WebkitAppearance: 'none',
        }}
        onFocus={e => e.target.style.outline = 'none'}
        onBlur={e => e.target.style.outline = 'none'}
        onKeyDown={e => { if (e.key === 'Enter') handleSend(); }}
      />
      <div style={{ display: 'flex', alignItems: 'center', gap: 32, marginTop: 0, paddingLeft: 0, paddingRight: 0 }}>
        <Button appearance="subtle" style={{ color: '#222', minWidth: 32, minHeight: 32, borderRadius: 6, padding: 0 }} icon={<Add24Regular style={{ fontSize: 20 }} />} aria-label="Add" />
      </div>
      {input ? (
        <Button
          appearance="primary"
          style={{
            position: 'absolute',
            right: 16,
            bottom: 8,
            minWidth: 32,
            minHeight: 32,
            width: 32,
            height: 32,
            borderRadius: '50%',
            background: '#2564cf',
            color: '#fff',
            fontWeight: 600,
            boxShadow: '0 1px 4px #0002',
            zIndex: 2,
            padding: 0,
            display: 'flex',
            alignItems: 'center',
            justifyContent: 'center',
          }}
          icon={<ArrowRight24Regular style={{ fontSize: 20 }} />}
          aria-label="Send"
          onClick={handleSend}
        />
      ) : (
        <Button
          appearance="subtle"
          style={{
            position: 'absolute',
            right: 16,
            bottom: 8,
            minWidth: 32,
            minHeight: 32,
            borderRadius: 8,
            color: '#222',
            background: 'transparent',
            boxShadow: 'none',
            padding: 0,
          }}
          icon={<Mic24Regular style={{ fontSize: 20 }} />}
          aria-label="Mic"
        />
      )}
    </div>
  );
}

export default App;
