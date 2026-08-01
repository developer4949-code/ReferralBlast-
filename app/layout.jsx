import './globals.css'

export const metadata = {
  title: 'ReferralBlast',
  description: 'Automate your referral emails with ease.',
}

export default function RootLayout({ children }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  )
}
