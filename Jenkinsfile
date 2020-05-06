pipeline {
   agent any
   
   tools {
        maven 'maven_3_6_3' 
    }
   stages {
      stage('Checking version') {
         steps {
            sh "mvn -version"
         }
    }
      stage('Compile stage') {
         steps {
            sh "mvn clean compile" 
        }
    }
      stage('Testing stage') {
         steps {
            sh "mvn test"
        }
    }
      stage('Deployment stage') {
         steps {
            echo "mvn deploy"
        }
    }  
   }
}
