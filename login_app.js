// Login authentication script
const SESSION_KEY = 'app_session_token'

// Valid credentials
const VALID_USERS = {
  'didin': '86532',
  'indra': '86086',
  'nur': '80229',
  'desi': '82847'
}

document.addEventListener('DOMContentLoaded', function(){
  const form = document.getElementById('login-form')
  const usernameInput = document.getElementById('username')
  const passwordInput = document.getElementById('password')
  const errorMessage = document.getElementById('error-message')
  
  form.addEventListener('submit', function(e){
    e.preventDefault()
    
    const username = usernameInput.value.trim()
    const password = passwordInput.value
    
    console.log('Login attempt:', {username, password})
    
    // Validate
    if(!username || !password){
      showError('Username dan sandi harus diisi')
      return
    }
    
    // Check credentials
    if(VALID_USERS[username] && VALID_USERS[username] === password){
      console.log('Login successful for user:', username)
      // Create session token
      const token = btoa(username + ':' + Date.now())
      localStorage.setItem(SESSION_KEY, token)
      localStorage.setItem('current_user', username)
      
      console.log('Session saved, redirecting...')
      // Redirect to layout
      window.location.href = 'layout.html'
    } else {
      console.log('Login failed - invalid credentials')
      showError('Username atau sandi salah')
      passwordInput.value = ''
      passwordInput.focus()
    }
  })
  
  function showError(message){
    errorMessage.textContent = message
    errorMessage.classList.add('show')
    setTimeout(() => {
      errorMessage.classList.remove('show')
    }, 4000)
  }
})
