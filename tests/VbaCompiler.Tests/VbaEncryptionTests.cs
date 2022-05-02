using System;

namespace VbaCompiler.Tests
{
    public class VbaEncryptionTests
    {
        [SetUp]
        public void SetUp()
        {
            VbaEncryption.CreateRandomGenerator = () => new VbaSamplesNumberGenerator();
        }

        [Test]
        [TestCase("{917DED54-440B-4FD1-A5C1-74ACF261E600}", (ushort)0xDF)]
        public void GetProjectKey_SampleProjectId_ComputesChecksumForEncryption(string projectId, ushort expectedProjectKey)
        {
            // Arrange

            // Act
            var actualProjectKey = VbaEncryption.GetProjectKey(projectId);

            // Assert
            Assert.AreEqual(expectedProjectKey, actualProjectKey);
        }

        [Test]
        [TestCase("{917DED54-440B-4FD1-A5C1-74ACF261E600}", "0705D8E3D8EDDBF1DBF1DBF1DBF1")]
        public void ProjectProtectionState_ToEncryptedString_Sample(string projectId, string expectedProtectionState)
        {
            // Arrange
            var state = new ProjectProtectionState(projectId);
            state.ProjectProtection = ProjectProtection.None;

            // Act
            var actualProtectionState = state.ToEncryptedString(0x07);

            // Assert
            Assert.AreEqual(expectedProtectionState, actualProtectionState);
        }

        [Test]
        [TestCase("{917DED54-440B-4FD1-A5C1-74ACF261E600}", "0E0CD1ECDFF4E7F5E7F5E7")]
        public void ProjectPassword_ToEncryptedString_Sample(string projectId, string expectedProtectionState)
        {
            // Arrange
            const byte seed = 0x0E;
            var state = new ProjectPassword(projectId);

            // Act
            var actualProtectionState = state.ToEncryptedString(seed);

            // Assert
            Assert.AreEqual(expectedProtectionState, actualProtectionState);
        }
    }
}