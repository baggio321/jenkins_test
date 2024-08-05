pipeline {
    agent any
    stages {
        stage('diff') {
            steps {
                bat 'python check.py'
            }
        }
    }
}
